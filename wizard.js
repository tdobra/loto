"use strict";

//TODO: Chrome has bug that defer script loading doesn't work with XHTML
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
      this.parentItem = obj.parentItem; //May be undefined
      this.itemInFocus = obj.itemInFocus === undefined ? 0 : obj.itemInFocus;
      this.items = [];
      this.counterField = obj.counterField;
    }

    get activeItem() {
      return courseList.viewDefault.value ? this.default : this.items[this.itemInFocus];
    }

    static init() {
      //Retrieve elements for given IDs and populate field names
      id2element(this, ["selector"].concat(this.btnList));
      this.resetGroups.forEach((group) => { id2element(group, ["resetBtn", "setAllBtn"]); });
      this.itemClass.init();

      //Add event listeners
      //Always use arrow functions to ensure this points to the IterableList rather than the calling DOM element
      //Course order buttons
      this.selector.addEventListener("change", () => { this.activeList.refreshItem(); });
      this.addBtn.addEventListener("click", () => {
        this.activeList.add();
        this.activeList.updateCount();
        this.activeList.refreshList();
      });
      this.deleteBtn.addEventListener("click", () => {
        this.activeList.activeItem.deleteThis(true);
        this.activeList.updateCount();
      });
      this.upBtn.addEventListener("click", () => { this.activeList.activeItem.move(-1); });
      this.downBtn.addEventListener("click", () => { this.activeList.activeItem.move(1); });
      //Fields
      this.itemClass.fieldClasses.forEach((field, index) => {
        const activeField = () => this.activeList.activeItem[this.itemClass.fieldNames[index]];
        if (field.inputElement.tagName === undefined) {
          //Radio buttons
          Object.values(field.inputElement).forEach((element) => {
            element.addEventListener("change", () => { activeField().saveInput(); });
          });
        } else {
          const inputEvent = (field.inputElement.tagName === "SELECT" || field.inputElement.type === "checkbox") ? "change" : "input";
          field.inputElement.addEventListener(inputEvent, () => { activeField().saveInput(); });
        }
        if (field.autoElement !== undefined) {
          field.autoElement.addEventListener("change", () => { activeField().auto.saveInput(); });
        }
        if (field.ruleElement !== undefined) {
          field.ruleElement.addEventListener("change", () => { activeField().ignoreRules.saveInput(); });
        }
        if (field.resetBtn !== undefined) {
          field.resetBtn.addEventListener("click", () => { activeField().resetValue(); });
        }
        if (field.setAllBtn !== undefined) {
          field.setAllBtn.addEventListener("click", () => { activeField().setAll(); });
        }
      });
      //Field groups
      this.resetGroups.forEach((group) => {
        group.resetBtn.addEventListener("click", () => { group.fieldClasses.forEach((field) => { field.resetBtn.click(); }); });
        group.setAllBtn.addEventListener("click", () => { group.fieldClasses.forEach((field) => { field.setAllBtn.click(); }); });
      });
    }

    static showHideMoveBtns() {
      //Show the delete button only if more than one option
      const numItems = this.selector.length;
      if (numItems > 1) {
        this.deleteBtn.disabled = false;
      } else {
        this.deleteBtn.disabled = true;
      }

      //Shows or hides the move up and move down buttons
      const index = this.selector.selectedIndex;
      if (index === 0) {
        //First station, so can't move it up
        this.upBtn.disabled = true;
      } else {
        this.upBtn.disabled = false;
      }
      if (index === numItems - 1) {
        //First station, so can't move it up
        this.downBtn.disabled = true;
      } else {
        this.downBtn.disabled = false;
      }
    }

    add() {
      //New item initially takes default values
      const newItem = this.newItem(false);

      //Name the new item
      newItem.itemName.value = (this.items.length + 1).toString();
      //Check unique and increment by 1 if not
      newItem.itemName.checkValidity();
      while (!newItem.itemName.valid) {
        newItem.itemName.value = (Number(newItem.itemName.value) + 1).toString();
        newItem.itemName.checkValidity();
      }
      newItem.setOptionText(newItem.itemName.value);
      newItem.checkValidity(true);

      //Add to the end and give focus
      this.itemInFocus = this.items.push(newItem) - 1;

      return newItem;
    }

    applyAll(action, fieldList) {
      //Applies specified action function to all fields with reset buttons
      fieldList.forEach((field) => {
        if (this.activeItem[field].resetBtn !== undefined) {
          this.activeItem[field][action]();
        }
      });
    }

    checkValidity(recheckFields = true) {
      if (recheckFields) { this.items.forEach((item) => { item.checkValidity(true); }); }
      this.valid = this.items.every((item) => item.valid);
    }

    newItem(isDefault = false) {
      return new this.constructor.itemClass({
        parentList: this,
        isDefault: isDefault,
        copy: (this.default === undefined) ? undefined : this.default
      });
    }

    refreshItem() {
      if (!courseList.viewDefault.value) {
        this.itemInFocus = this.constructor.selector.selectedIndex;
        this.constructor.showHideMoveBtns();
      }
      this.activeItem.refreshAllInput(true);
    }

    refreshList() {
      //Remove all nodes
      while (this.constructor.selector.hasChildNodes()) {
        this.constructor.selector.removeChild(this.constructor.selector.lastChild);
      }
      //Add nodes for current item
      this.items.forEach((item) => { this.constructor.selector.appendChild(item.optionElement); });
      this.constructor.selector.selectedIndex = this.itemInFocus;
      //Refresh fields
      this.refreshItem();
    }

    updateCount() {
      if (this.counterField !== undefined) {
        this.counterField.save(this.items.length);
        this.counterField.refreshInput(); //FIXME: parameters to function
      }
    }
  }
  Object.assign(IterableList, {
    btnList: ["addBtn", "deleteBtn", "upBtn", "downBtn"],
    resetGroups: []
  });

  class IterableItem {
    constructor(obj) {
      this.parentList = obj.parentList;
      this.optionCopies = [];
      if (!obj.isDefault) {
        this.optionElement = document.createElement("option");
        for (let i = 0; i < obj.numOptionCopies; ++i) { this.optionCopies.push(this.optionElement.cloneNode(false)); }
      }
      this.createFields(obj.copy);
    }

    get index() {
      let thisIndex = 0;
      while (this.parentList.items[thisIndex] !== this) { ++thisIndex; }
      return thisIndex;
    }

    static init() {
      this.fieldNames = this.fieldClasses.map((className) => {
        className.init();
        return className.fieldName;
      });
    }

    isActive() {
      return this.parentList.activeItem === this;
    }

    isDefault() {
      return this.parentList.default === this;
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
      this.parentList.constructor.selector.selectedIndex = 0;
      //Remove this station from array
      this.parentList.items.splice(this.index, 1);
      this.parentList.refreshItem();
    }

    move(offset = 0) {
      //Moves this item
      paramsSaved = false;

      //Check target is in bounds
      const newPos = this.index + offset;
      const arrayLength = this.parentList.items.length;
      if (newPos < 0 || newPos >= arrayLength) {
        throw new Error("Moving station to position beyond bounds of array");
      }
      if (!Number.isInteger(newPos)) {
        throw new Error("Station position offset is not an integer");
      }

      //Remove this from array
      this.parentList.items.splice(this.index, 1);
      //Insert into new position
      this.parentList.items.splice(newPos, 0, this);

      //Rearrange selector
      let insertBeforeElement = null;
      if (newPos < arrayLength - 1) {
        insertBeforeElement = this.parentList.items[newPos + 1].optionElement;
      } //Else move to end of list
      //Make sure to get a new element list, as it will have changed
      this.parentList.constructor.selector.insertBefore(this.optionElement, insertBeforeElement);

      //Update active station number
      this.parentList.itemInFocus = newPos;

      //Enable/disable move up/down buttons
      this.parentList.constructor.showHideMoveBtns();
    }

    refreshAllInput(inputChanged = true) {
      this.constructor.fieldNames.forEach((field) => { this[field].refreshInput(true, inputChanged); });
    }

    setOptionText(val) {
      this.optionElement.textContent = val;
      this.optionCopies.forEach((el) => { el.textContent = val; });
    }
  }

  class Field {
    //Generic field: string or select
    constructor(obj) {
      this.parentItem = obj.parentItem;
      if (this.parentItem.constructor.itemClass === undefined) {
        //parentItem is an IterableItem
        this.parentList = this.parentItem.parentList;
        if (this.parentList.constructor.itemClass === Course || this.parentList.constructor.itemClass === Station) {
          this.topItem = this.parentItem;
        } else {
          this.topItem = this.parentList.parentItem;
        }
        this.topList = this.topItem.parentList;
      }
      this.inputElement = obj.inputElement === undefined ? this.constructor.inputElement : obj.inputElement;
      this.value = obj.value === undefined ? this.constructor.originalValue : obj.value;
      this.valid = true;
      this.ruleCompliant = true;
      if (this.constructor.autoElement !== undefined) {
        this.auto = new AutoField({
          parentItem: obj.parentItem,
          parentField: this,
          inputElement: this.constructor.autoElement
        });
      }
      if (this.constructor.ruleElement !== undefined) {
        this.ignoreRules = new IgnoreRules({
          parentItem: obj.parentItem,
          parentField: this,
          inputElement: this.constructor.ruleElement
        });
      }
    }

    get inputValue() {
      return this.inputElement.value;
    }

    set inputValue(val) {
      this.inputElement.value = val;
    }

    static init() {
      id2element(this, ["inputElement", "autoElement", "ruleElement", "resetBtn", "setAllBtn", "errorElement"]);
    }

    autoEnabled() {
      return (this.auto !== undefined) ? this.auto.value : false;
    }

    isDuplicate(ignoreHidden = false) {
      //Determines whether this value is duplicated at another item in this list. Ignores hidden stations on request.

      //Return false if the tested index is ignored
      if (ignoreHidden && this.topItem.isHidden()) { return false; }

      const isDup = (item) => item[this.constructor.fieldName].value === this.value && item !== this.parentItem;

      if (ignoreHidden) {
        //FIXME: ITEM.ISHIDDEN probably doesn't work for kites/tasks
        return this.parentList.items.some((item) => isDup(item) || item.isHidden());
      } else {
        return this.parentList.items.some(isDup);
      }
    }

    matchesAll(ignoreHidden = false, skipIgnoreRules = false) {
      //Returns true if this field matches those on all other items in this list, excluding number fields set to NaN. Ignores NaN if number. Ignores hidden stations on request.

      //Return true if the tested index is ignored
      if (ignoreHidden && this.topItem.isHidden()) { return true; }
      if (skipIgnoreRules) { if (this.ignoreRules.value) { return true; } }

      const isMatch = (item) => item[this.constructor.fieldName].value === this.value || Number.isNaN(item[this.constructor.fieldName].value);
      //FIXME: item.isHidden() doesn't work for kites/tasks
      const isMatchHidden = (item) => isMatch(item) || (ignoreHidden && item.isHidden());

      let isMatchRules;
      if (skipIgnoreRules) {
        //Only call if ignoreRules is defined
        isMatchRules = (item) => isMatchHidden(item) || item[this.constructor.fieldName].ignoreRules.value;
      } else {
        isMatchRules = isMatchHidden;
      }
      return this.parentList.items.every(isMatchRules);
    }

    refreshInput(alsoSubfields = false, inputChanged = true) {
      //alsoSubfields false avoids infinite recursion
      //inputChanged only accesses DOM if required
      //Updates input element value - no user input so no change in validity
      if (alsoSubfields) {
        ["auto", "ignoreRules"].forEach((fieldName) => {
          if (this[fieldName] !== undefined) { this[fieldName].refreshInput(false, inputChanged); }
        });
      }
      if (inputChanged) { this.inputValue = this.value; }
      this.updateMsgs();
    }

    resetValue() {
      const resetField = (field) => {
        let val;
        if (this.topItem.isDefault()) {
          //Resets to original value
          val = field.constructor.originalValue;
        } else {
          //Resets to current default value
          let defaultItem;
          switch (this.parentList.itemClass) {
            case Kite:
              defaultItem = this.topList.default.kites.item[0];
              break;
            case Task:
              defaultItem = this.topList.default.tasks.item[0];
              break;
            default:
              defaultItem = this.topList.default;
          }
          let defaultField = defaultItem[this.constructor.fieldName];
          switch (field.constructor.fieldName) {
            case "auto":
            case "ignoreRules":
              defaultField = defaultField[field.constructor.fieldName];
          }
          val = defaultField.value;
        }
        field.save(val);
      };

      if (this.ignoreRules !== undefined) { resetField(this.ignoreRules); }
      if (this.auto !== undefined) {
        resetField(this.auto);
        if (!this.auto.value) { resetField(this); }
      } else {
        resetField(this);
      }
      this.refreshInput(true, true);
    }

    saveInput() {
      //Call when the value of the input element is updated by the user
      this.save(this.inputValue);
      this.updateMsgs();
    }

    save(val) {
      //Saves value and determines whether a change has occurred
      if (this.value !== val) { this.saveValue(val); }
    }

    saveValue(val) {
      //Flag value as changed
      paramsSaved = false;
      //Write new value to memory
      this.value = val;
      //Check whether this new value is valid
      let items;
      if (this.constructor.checkSiblings) {
        //Also check siblings for matching errors
        items = this.parentList.items.map((item) => item[this.constructor.fieldName]);
        items.push(this);
      } else {
        items = [this];
      }
      items.forEach((item) => {
        item.checkValidity();
        item.parentItem.checkValidity(false);
      });
    }

    setAll() {
      //Sets this field to this value in all items
      const setAllList = (list) => {
        list.items.forEach((item) => {
          const field = item[this.constructor.fieldName];
          if (field.ignoreRules !== undefined) { field.ignoreRules.save(this.ignoreRules.value); }
          if (field.auto !== undefined) {
            field.auto.save(this.auto.value);
            if (!field.auto.value) { field.save(this.value); }
          } else {
            field.save(this.value);
          }
        });
      }

      if (this.parentList.itemClass === Kite) {
        //Apply to all stations
        this.topList.items.forEach((item) => { setAllList(item.kites); });
      } else if (this.parentList.itemClass === Task) {
        this.topList.items.forEach((item) => { setAllList(item.tasks); });
      } else {
        setAllList(this.parentList);
      }
    }

    updateMsgs() {
      const contentFieldClass = this.inputElement.classList;
      const autoEnabled = this.auto !== undefined && this.auto.value;
      const errorMsg = autoEnabled ? tcTemplateMsg.awaitingData : this.errorMsg;
      if (this.valid || (this.topItem.isDefault() && autoEnabled)) {
        //Don't flag error on defaults if setting automatically
        contentFieldClass.remove("error");
        if (this.ruleCompliant) {
          contentFieldClass.remove("warning");
          if (this.constructor.errorElement !== undefined) {
            this.constructor.errorElement.textContent = "";
          }
        } else {
          contentFieldClass.add("warning");
          this.constructor.errorElement.textContent = errorMsg;
        }
      } else {
        contentFieldClass.remove("warning");
        contentFieldClass.add("error");
        this.constructor.errorElement.textContent = errorMsg;
      }
    }

    //Empty functions, which may be overwritten in inheriting classes if some action is required
    checkValidity() { } //Most likely: field always valid
  }
  Object.assign(Field, {
    originalValue: "",
    checkSiblings: false
  });

  class BooleanField extends Field {
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

    saveValue(val) {
      if (val) {
        //Overwrite parent field value with the calculated value
        this.parentField.saveValue(this.calculated);
        this.parentField.checkValidity();
      }
      super.saveValue(val);
    }

    updateMsgs() {
      if (this.value) {
        this.parentField.inputElement.readOnly = true;
        this.parentField.refreshInput(false, true);
      } else {
        this.parentField.inputElement.readOnly = false;
      }
      this.parentField.checkValidity();
      this.parentField.updateMsgs();
    }
  }
  Object.assign(AutoField, {
    fieldName: "auto",
    originalValue: true
  });

  class IgnoreRules extends BooleanField {
    constructor(obj) {
      super(obj);
      this.parentField = obj.parentField;
    }

    updateMsgs() {
      this.parentField.checkValidity();
      this.parentField.refreshInput(false, false);
    }
  }
  IgnoreRules.fieldName = "ignoreRules";

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
      this.valid = Number.isFinite(this.value) || this.topItem.isNonDefaultHidden();
    }

    save(val) {
      //New & saved values not equal and at least one of them not NaN (NaN === NaN is false)
      if (this.value !== val && (Number.isNaN(this.value) === false || Number.isNaN(val) === false)) { this.saveValue(val); }
    }
  }
  NumberField.originalValue = NaN;

  class NonNegativeField extends NumberField {
    constructor(obj) {
      super(obj);
      this.errorMsg = tcTemplateMsg.nonNegativeFieldError;
    }

    checkValidity() {
      super.checkValidity();
      this.valid = this.valid && this.value >= 0;
    }
  }

  class StrictPositiveField extends NonNegativeField {
    constructor(obj) {
      super(obj);
      this.errorMsg = tcTemplateMsg.strictPositiveFieldError;
    }

    checkValidity() {
      //LaTeX crashes if font size is set to zero
      super.checkValidity();
      this.valid = this.valid && this.value > 0;
    }
  }

  class NaturalNumberField extends NonNegativeField {
    constructor(obj) {
      super(obj);
      this.errorMsg = tcTemplateMsg.naturalNumberFieldError;
    }

    checkValidity() {
      this.valid = (Number.isInteger(this.value) && this.value > 0) || this.topItem.isNonDefaultHidden();
    }
  }

  class RadioField extends Field {
    static init() {
      super.init();
      id2element(this.inputElement, Object.keys(this.inputElement));
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

    updateMsgs() { }
  }

  class NameField extends Field {
    checkValidity() {
      if (!this.topItem.isDefault()) {
        //Check syntax even if hidden to avoid dodgy strings getting into LaTeX
        const stringFormat = /^[A-Za-z0-9][,.\-+= \w]*$/;
        const syntaxError = !stringFormat.test(this.value);
        const duplicateError = this.isDuplicate(false);
        this.valid = !(syntaxError || duplicateError);
        if (!this.valid) { this.errorMsg = syntaxError ? tcTemplateMsg.nameSyntax : tcTemplateMsg.notUnique; }
        //Update station list
        this.parentItem.setOptionText(this.value);
      } //Else remains valid
    }
  }
  Object.assign(NameField, {
    fieldName: "itemName",
    checkSiblings: true
  });

  //Course - in reverse order of dependency
  class CourseName extends NameField { }
  Object.assign(CourseName, {
    inputElement: "courseName",
    errorElement: "courseNameError"
  });

  class TasksFile extends RadioField {
    saveInput() {
      super.saveInput();
      this.parentItem.checkValidity(true);
      this.parentItem.refreshAllInput(false);
    }
  }
  Object.assign(TasksFile, {
    fieldName: "tasksFile",
    originalValue: "newFile",
    inputElement: {
      newFile: "newTasksFile",
      append: "appendTasksFile",
      hide: "hideTasksFile"
    },
    resetBtn: "resetTasksFile",
    setAllBtn: "setAllTasksFile"
  });

  class TasksTemplate extends Field { }
  Object.assign(TasksTemplate, {
    fieldName: "tasksTemplate",
    originalValue: "printA5onA4",
    inputElement: "tasksTemplate"
  });

  class AppendTasksCourse extends Field { }
  Object.assign(AppendTasksCourse, {
    fieldName: "appendTasksCourse",
    inputElement: "appendTasksCourse"
  });

  class Zeroes extends BooleanField {
    updateMsgs() {
      dynamicCSS.hide(!this.value, "#mainView .zeroes");
    }
  }
  Object.assign(Zeroes, {
    fieldName: "zeroes",
    inputElement: "zeroes",
    resetBtn: "resetZeroes",
    setAllBtn: "setAllZeroes"
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
    inputElement: "numTasks",
    resetBtn: "resetNumTasks",
    setAllBtn: "setAllNumTasks",
    errorElement: "numTasksError"
  });

  class MapScale extends StrictPositiveField {
    checkValidity() {
      super.checkValidity();
      if (this.valid) {
        //Check all values match
        this.ruleCompliant = this.parentItem.isDefault() || this.matchesAll(true, !this.ignoreRules.value);
        if (this.ruleCompliant) {
          //Permitted map scales is should, so don't throw error
          this.ruleCompliant = this.value === 4000 || this.value === 5000;
          this.errorMsg = tcTemplateMsg.mapScaleRule;
        } else {
          this.errorMsg = tcTemplateMsg.notSameForAllCourses;
          this.valid = this.ignoreRules.value;
        }
      } else {
        this.errorMsg = tcTemplateMsg.strictPositiveFieldError;
      }
    }
  }
  Object.assign(MapScale, {
    fieldName: "mapScale",
    originalValue: 4000,
    checkSiblings: true,
    inputElement: "mapScale",
    ruleElement: "mapScaleRule",
    resetBtn: "resetMapScale",
    setAllBtn: "setAllMapScale",
    errorElement: "mapScaleError",
  });

  class ContourInterval extends StrictPositiveField {
    checkValidity() {
      super.checkValidity();
      if (this.valid) {
        //Check all values match
        this.ruleCompliant = this.parentItem.isDefault() || this.matchesAll(true, !this.ignoreRules.value);
        this.valid = this.ignoreRules.value || this.ruleCompliant;
        this.errorMsg = tcTemplateMsg.notSameForAllCourses;
      } else {
        this.errorMsg = tcTemplateMsg.strictPositiveFieldError;
      }
    }
  }
  Object.assign(ContourInterval, {
    fieldName: "contourInterval",
    checkSiblings: true,
    inputElement: "contourInterval",
    ruleElement: "contourIntervalRule",
    resetBtn: "resetContourInterval",
    setAllBtn: "setAllContourInterval",
    errorElement: "contourIntervalError"
  });

  class MapShape extends Field {
    constructor(obj) {
      super(obj);
      this.errorMsg = tcTemplateMsg.notSameForAllCourses;
    }

    static init() {
      super.init();
      id2element(this, ["sizeTypeElement"]);
    }

    checkValidity() {
      //Don't throw error for not matching
      this.ruleCompliant = this.parentItem.isDefault() || this.matchesAll(true, false);
    }

    updateMsgs() {
      super.updateMsgs();
      //Update labels for map size description
      if (this.value === "circle") {
        this.constructor.sizeTypeElement.textContent = tcTemplateMsg.diameter;
      } else {
        this.constructor.sizeTypeElement.textContent = tcTemplateMsg.sideLength;
      }
    }
  }
  Object.assign(MapShape, {
    fieldName: "mapShape",
    originalValue: "circle",
    checkSiblings: true,
    inputElement: "mapShape",
    ruleElement: "mapShapeRule",
    resetBtn: "resetMapShape",
    setAllBtn: "setAllMapShape",
    errorElement: "mapShapeError",
    sizeTypeElement: "mapSizeType"
  });

  class MapSize extends StrictPositiveField {
    checkValidity() {
      super.checkValidity();
      this.valid = this.valid && this.value <= 12;
      if (this.valid) {
        this.ruleCompliant = this.value >= 5;
        if (this.ruleCompliant) {
          //Don't throw error for not matching
          this.ruleCompliant = this.parentItem.isDefault() || this.matchesAll(true, false);
          this.errorMsg = tcTemplateMsg.notSameForAllCourses;
        } else {
          this.errorMsg = tcTemplateMsg.mapSizeRule;
          this.valid = this.ignoreRules.value;
        }
      } else {
        this.errorMsg = tcTemplateMsg.mapSizeError;
      }
    }
  }
  Object.assign(MapSize, {
    fieldName: "mapSize",
    checkSiblings: true,
    inputElement: "mapSize",
    autoElement: "mapSizeAuto",
    ruleElement: "mapSizeRule",
    resetBtn: "resetMapSize",
    setAllBtn: "setAllMapSize",
    errorElement: "mapSizeError"
  });

  class NumKites extends NumberField {
    constructor(obj) {
      super(obj);
      this.errorMsg = tcTemplateMsg.numKitesRule;
    }

    checkValidity() {
      this.ruleCompliant = this.topItem.isHidden() || this.value === 6;
      this.valid = this.ignoreRules.value || this.ruleCompliant;
    }
  }
  Object.assign(NumKites, {
    fieldName: "numKites",
    originalValue: 6,
    inputElement: "numKites",
    ruleElement: "numKitesRule",
    resetBtn: "resetNumKites",
    setAllBtn: "setAllNumKites",
    errorElement: "numKitesError"
  });

  class DebugCircle extends BooleanField { }
  Object.assign(DebugCircle, {
    fieldName: "debugCircle",
    inputElement: "debugCircle",
    resetBtn: "resetDebugCircle",
    setAllBtn: "setAllDebugCircle"
  });

  class IDFontSize extends StrictPositiveField { }
  Object.assign(IDFontSize, {
    fieldName: "IDFontSize",
    originalValue: 0.7,
    inputElement: "IDFontSize",
    resetBtn: "resetIDFontSize",
    setAllBtn: "setAllIDFontSize",
    errorElement: "IDFontSizeError"
  });

  class CheckWidth extends NonNegativeField { }
  Object.assign(CheckWidth, {
    fieldName: "checkWidth",
    originalValue: 1.5,
    inputElement: "checkWidth",
    resetBtn: "resetCheckWidth",
    setAllBtn: "setAllCheckWidth",
    errorElement: "checkWidthError"
  });

  class CheckHeight extends NonNegativeField { }
  Object.assign(CheckHeight, {
    fieldName: "checkHeight",
    originalValue: 1.5,
    inputElement: "checkHeight",
    resetBtn: "resetCheckHeight",
    setAllBtn: "setAllCheckHeight",
    errorElement: "checkHeightError"
  });

  class CheckFontSize extends StrictPositiveField { }
  Object.assign(CheckFontSize, {
    fieldName: "checkFontSize",
    originalValue: 0.8,
    inputElement: "checkFontSize",
    resetBtn: "resetCheckFontSize",
    setAllBtn: "setAllCheckFontSize",
    errorElement: "checkFontSizeError"
  });

  class RemoveFontSize extends StrictPositiveField { }
  Object.assign(RemoveFontSize, {
    fieldName: "removeFontSize",
    originalValue: 0.3,
    inputElement: "removeFontSize",
    resetBtn: "resetRemoveFontSize",
    setAllBtn: "setAllRemoveFontSize",
    errorElement: "removeFontSizeError"
  });

  class PointHeight extends NonNegativeField { }
  Object.assign(PointHeight, {
    fieldName: "pointHeight",
    originalValue: 2.5,
    inputElement: "pointHeight",
    resetBtn: "resetPointHeight",
    setAllBtn: "setAllPointHeight",
    errorElement: "pointHeightError"
  });

  class LetterFontSize extends StrictPositiveField { }
  Object.assign(LetterFontSize, {
    fieldName: "letterFontSize",
    originalValue: 1.8,
    inputElement: "letterFontSize",
    resetBtn: "resetLetterFontSize",
    setAllBtn: "setAllLetterFontSize",
    errorElement: "letterFontSizeError"
  });

  class PhoneticFontSize extends StrictPositiveField { }
  Object.assign(PhoneticFontSize, {
    fieldName: "phoneticFontSize",
    originalValue: 0.6,
    inputElement: "phoneticFontSize",
    resetBtn: "resetPhoneticFontSize",
    setAllBtn: "setAllPhoneticFontSize",
    errorElement: "phoneticFontSizeError"
  });

  class Course extends IterableItem {
    constructor(obj) {
      //Option copy allocation: 0 = StationCourse
      obj.numOptionCopies = 1;
      super(obj);
    }

    isHidden() {
      return this.tasksFile.value === "hide";
    }

    deleteThis() {
      super.deleteThis();
      this.parentList.refreshOtherCourseLists();
      //Change any stationCourse entries that point to this deleted station
      if (stationList.default.stationCourse === this) { stationList.default.stationCourse = StationCourse.originalValue; }
      stationList.items.forEach((station) => {
        if (station.stationCourse === this) { station.stationCourse = stationList.default.stationCourse; }
      });
    }

    move(offset) {
      super.move(offset);
      this.parentList.refreshOtherCourseLists();
    }
  }
  Course.customLayoutClasses = [
    DebugCircle,
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
    NumKites
  ].concat(Course.customLayoutClasses);

  class ViewDefault extends BooleanField {
    saveInput() {
      super.saveInput();
      this.parentItem.refreshItem();
    }

    updateMsgs() {
      //Show/hide any buttons as required
      dynamicCSS.hide(this.value, "#mainView .hideIfDefault");
      dynamicCSS.hide(!this.value, ".setAllCourses");
    }
  }
  Object.assign(ViewDefault, {
    inputElement: "viewDefault"
  });

  class CourseList extends IterableList {
    constructor() {
      const obj = {
        itemInFocus: 0,
        counterField: undefined
      };
      super(obj);
      //Set up current/dynamic defaults
      this.viewDefault = new ViewDefault({ parentItem: this });
      this.default = this.newItem(true);
      this.default.checkValidity(true);
    }

    static get activeList() {
      return courseList;
    }

    static init() {
      super.init();
      id2element(ViewDefault, ["inputElement"]);
      ViewDefault.inputElement.addEventListener("change", () => { courseList.viewDefault.saveInput(); });
    }

    add() {
      super.add();
      this.refreshOtherCourseLists();
    }

    refreshItem() {
      this.viewDefault.refreshInput(false, true);
      super.refreshItem();
    }

    refreshOtherCourseLists() {
      StationCourse.refreshList();
    }
  }
  Object.assign(CourseList, {
    selector: "courseSelect",
    addBtn: "addCourse",
    deleteBtn: "deleteCourse",
    upBtn: "moveUpCourse",
    downBtn: "moveDownCourse",
    itemClass: Course,
    resetGroups: [
      {
        resetBtn: "resetAllCustomLayout",
        setAllBtn: "setAllCustomLayout",
        fieldClasses: Course.customLayoutClasses
      }
    ]
  });

  //Station - in reverse order of dependency
  class StationName extends NameField { }
  Object.assign(StationName, {
    inputElement: "stationName",
    errorElement: "stationNameError"
  });

  class StationCourse extends Field {
    static get originalValue() {
      return courseList.items[courseList.itemInFocus];
    }

    get inputValue() {
      return courseList.items[this.inputElement.selectedIndex];
    }

    set inputValue(val) {
      this.inputElement.value = val.itemName.value;
    }

    static refreshList() {
      //Remove existing contents
      while (this.inputElement.hasChildNodes()) { this.inputElement.removeChild(this.inputElement.lastChild); }
      //Repopulate
      courseList.items.forEach((course) => { this.inputElement.appendChild(course.optionCopies[0]); });
    }

    updateMsgs() {
      super.updateMsgs();
      //Show relevant course
      CourseList.selector.selectedIndex = this.value.index;
      courseList.refreshItem();
    }
  }
  Object.assign(StationCourse, {
    fieldName: "stationCourse",
    inputElement: "stationCourse",
    resetBtn: "resetStationCourse",
    setAllBtn: "setAllStationCourse"
  });

  class ShowStationTasks extends BooleanField {
    saveInput() {
      super.saveInput();
      this.parentItem.checkValidity(true);
      this.parentItem.refreshAllInput(false);
    }
  }
  Object.assign(ShowStationTasks, {
    fieldName: "showStationTasks",
    originalValue: true,
    inputElement: "showStationTasks",
    resetBtn: "resetShowStationTasks",
    setAllBtn: "setAllShowStationTasks"
  });

  class AutoPopulateKites extends BooleanField { }
  Object.assign(AutoPopulateKites, {
    fieldName: "autoPopulateKites",
    originalValue: true,
    inputElement: "kitesAutoPopulate",
    resetBtn: "resetKitesAutoPopulate",
    setAllBtn: "setAllKitesAutoPopulate"
  });

  class AutoOrderKites extends BooleanField { }
  Object.assign(AutoOrderKites, {
    fieldName: "autoOrderKites",
    originalValue: true,
    inputElement: "autoOrderKites",
    resetBtn: "resetAutoOrderKites",
    setAllBtn: "setAllAutoOrderKites"
  });

  class KiteName extends NameField { }
  Object.assign(KiteName, {
    inputElement: "kiteName",
    errorElement: "kiteNameError"
  });

  class KiteZero extends BooleanField { }
  Object.assign(KiteZero, {
    fieldName: "kiteZero",
    inputElement: "zero"
  });

  class Kitex extends NonNegativeField { }
  Object.assign(Kitex, {
    fieldName: "kitex",
    inputElement: "kitex",
    autoElement: "kitexAuto",
    errorElement: "kitexError"
  });

  class Kitey extends NonNegativeField { }
  Object.assign(Kitey, {
    fieldName: "kitey",
    inputElement: "kitey",
    autoElement: "kiteyAuto",
    errorElement: "kiteyError"
  });

  class Kite extends IterableItem { }
  Kite.fieldClasses = [
    KiteName,
    KiteZero,
    Kitex,
    Kitey
  ];

  class KiteList extends IterableList {
    static get activeList() {
      return stationList.activeItem.kites;
    }
  }
  Object.assign(KiteList, {
    selector: "kiteSelect",
    addBtn: "addKite",
    deleteBtn: "deleteKite",
    upBtn: "moveUpKite",
    downBtn: "moveDownKite",
    itemClass: Kite
  });

  class VPx extends NonNegativeField { }
  Object.assign(VPx, {
    fieldName: "vpx",
    inputElement: "VPx",
    autoElement: "VPxAuto",
    errorElement: "VPxError"
  });

  class VPy extends NonNegativeField { }
  Object.assign(VPy, {
    fieldName: "vpy",
    inputElement: "VPy",
    autoElement: "VPyAuto",
    errorElement: "VPyError"
  });

  class Heading extends NumberField { }
  Object.assign(Heading, {
    fieldName: "heading",
    inputElement: "heading",
    autoElement: "headingAuto",
    errorElement: "headingError"
  });

  class BlankMapPage extends NaturalNumberField { }
  Object.assign(BlankMapPage, {
    fieldName: "blankMapPage",
    inputElement: "blankMapPage",
    autoElement: "blankMapPageAuto",
    resetBtn: "resetBlankMapPage",
    setAllBtn: "setAllBlankMapPage",
    errorElement: "blankMapPageError"
  });

  class TaskName extends NameField {
    checkValidity() {
      if (!this.topItem.isDefault()) {
        super.checkValidity();
        if (
          this.valid && !this.topItem.isDefault() &&
          (this.value === "Kites" || this.value === "VP" || this.value === "Zs")
        ) {
          this.valid = false;
          this.errorMsg = tcTemplateMsg.taskNameKeywords;
        }
      }
    }
  }
  Object.assign(TaskName, {
    inputElement: "taskName",
    errorElement: "taskNameError"
  });

  class Solution extends Field { }
  Object.assign(Solution, {
    fieldName: "solution",
    inputElement: "solution",
    autoElement: "solutionAuto",
    errorElement: "solutionStatus"
  });

  class CirclePage extends NaturalNumberField { }
  Object.assign(CirclePage, {
    fieldName: "circlePage",
    inputElement: "circlePage",
    autoElement: "circlePageAuto",
    errorElement: "circlePageError"
  });

  class Circlex extends NonNegativeField { }
  Object.assign(Circlex, {
    fieldName: "circlex",
    inputElement: "circlex",
    autoElement: "circlexAuto",
    errorElement: "circlexError"
  });

  class Circley extends NonNegativeField { }
  Object.assign(Circley, {
    fieldName: "circley",
    inputElement: "circley",
    autoElement: "circleyAuto",
    errorElement: "circleyError"
  });

  class CDPage extends NaturalNumberField { }
  Object.assign(CDPage, {
    fieldName: "cdPage",
    inputElement: "CDPage",
    autoElement: "CDPageAuto",
    errorElement: "CDPageError"
  });

  class CDx extends NonNegativeField { }
  Object.assign(CDx, {
    fieldName: "cdx",
    inputElement: "CDx",
    autoElement: "CDxAuto",
    errorElement: "CDxError"
  });

  class CDy extends NonNegativeField { }
  Object.assign(CDy, {
    fieldName: "cdy",
    inputElement: "CDy",
    autoElement: "CDyAuto",
    errorElement: "CDyError"
  });

  class CDWidth extends NonNegativeField { }
  Object.assign(CDWidth, {
    fieldName: "cdWidth",
    inputElement: "CDWidth",
    autoElement: "CDWidthAuto",
    errorElement: "CDWidthError"
  });

  class CDHeight extends NonNegativeField { }
  Object.assign(CDHeight, {
    fieldName: "cdHeight",
    inputElement: "CDHeight",
    autoElement: "CDHeightAuto",
    errorElement: "CDHeightError"
  });

  class CDScale extends NonNegativeField { }
  Object.assign(CDScale, {
    fieldName: "cdScale",
    inputElement: "CDScale",
    autoElement: "CDScaleAuto",
    errorElement: "CDScaleError"
  });

  class Task extends IterableItem { }
  Task.fieldClasses = [
    TaskName,
    Solution,
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

  class TaskList extends IterableList {
    constructor(obj) {
      Object.assign(obj, {
        itemInFocus: 0,
        counterField: obj.parentItem.numTasks
      });
      super(obj);
    }

    static get activeList() {
      return stationList.activeItem.tasks;
    }

    refreshItem(storeValues = false) {
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

      this.itemInFocus = this.constructor.selector.selectedIndex;
      this.constructor.showHideMoveBtns();
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
  Object.assign(TaskList, {
    selector: "taskSelect",
    addBtn: "addTask",
    deleteBtn: "deleteTask",
    upBtn: "moveUpTask",
    downBtn: "moveDownTask",
    itemClass: Task
  });

  class Station extends IterableItem {
    constructor(obj) {
      super(obj);
      this.kites = new KiteList({ parentItem: this });
      this.tasks = new TaskList({ parentItem: this });

      //WARNING: need to make a copy of any sub objects, otherwise will still refer to same memory
      // this.taskList = new TaskList(this);
      // this.taskList.setNumItems();
    }

    isHidden() {
      return !this.showStationTasks.value || this.stationCourse.value.isHidden();
    }
  }
  Station.fieldClasses = [
    StationName,
    StationCourse,
    ShowStationTasks,
    AutoPopulateKites,
    AutoOrderKites,
    VPx,
    VPy,
    Heading,
    BlankMapPage
  ];

  class StationList extends IterableList {
    constructor() {
      const obj = {
        itemInFocus: 0,
        counterField: undefined
      };
      super(obj);
      //Set up current/dynamic defaults
      this.default = this.newItem(true);
      this.default.kites.add();
      this.default.tasks.add();
      this.default.checkValidity(true);
    }

    static get activeList() {
      return stationList;
    }

    add() {
      //Never called for default station
      const station = super.add();
      const course = station.stationCourse.value;
      for (let i = 0; i < course.numKites.value; ++i) { station.kites.add(); }
      for (let i = 0; i < course.numTasks.value; ++i) { station.tasks.add(); }
    }

    refreshItem() {
      super.refreshItem();
      //Clear and repopulate selectors
      [this.activeItem.kites, this.activeItem.tasks].forEach((list) => { list.refreshList(); });
    }
  }
  Object.assign(StationList, {
    selector: "stationSelect",
    addBtn: "addStation",
    deleteBtn: "deleteStation",
    upBtn: "moveUpStation",
    downBtn: "moveDownStation",
    itemClass: Station
  });

  function id2element(obj, props) {
    //Replaces properties on object that specify an element ID by the element itself
    props.forEach((elName) => { if (typeof obj[elName] === "string") { obj[elName] = document.getElementById(obj[elName]); } });
  }

  //Save and retrieve parameters
  function buildXML() {
    function createXMLFields(itemClass, item, parentNode) {
      itemClass.fieldClasses.forEach((field) => {
        const fieldName = field.fieldName;
        const newNode = xmlDoc.createElementNS(tctNamespace, fieldName);
        newNode.appendChild(xmlDoc.createTextNode(item[fieldName].value));
        if (field.autoElement !== undefined) { newNode.setAttribute("auto", item[fieldName].auto.value); }
        if (field.ruleElement !== undefined) { newNode.setAttribute("ignoreRules", item[fieldName].ignoreRules.value); }
        xmlNode.appendChild(newNode);
      });
    }

    const tctNamespace = "https://tdobra.github.io/tctemplate";
    const xmlDoc = document.implementation.createDocument(tctNamespace, "tctemplate");
    let xmlParent = xmlDoc.documentElement;

    let xmlNode = xmlDoc.createElementNS(tctNamespace, "default");
    createXMLFields(Course, courseList.default, xmlNode);
    createXMLFields(Station, stationList.default, xmlNode);
    // createXMLFields(Task, stationList.default.tasks.items[0], xmlNode);
    xmlParent.appendChild(xmlNode);


    //Convert to string
    const xmlSerial = new XMLSerializer();
    let xmlStr = xmlSerial.serializeToString(xmlDoc);
    //Add XML declaration if not already present - varies by browser
    if (!xmlStr.startsWith("<?")) {
      xmlStr = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + xmlStr;
    }
    console.log(xmlStr);
  }
  document.getElementById("saveParameters").addEventListener("click", buildXML);



















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
        courseOrderUsed.sort(function (a, b) { return a - b });

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
        selectors.push(selector);
      }
      return stylesheet.cssRules[index].style;
    }

    function hide(hide, selector) {
      const style = getStyle(selector);
      style.display = hide ? "none" : "";
    }

    return {
      hide: hide
    }
  })();

  [
    CourseList,
    StationList,
    KiteList,
    TaskList
  ].forEach((list) => { list.init(); });

  //Group auto/manual buttons
  class AutoBtnSet {
    //Set of buttons to set given fields of all items of given type in this/all station to automatic or manual
    constructor(obj) {
      ["listType", "fields", "autoAllBtn", "autoOneBtn", "manualAllBtn", "manualOneBtn"].forEach((fieldName) => {
        this[fieldName] = obj[fieldName];
      })
    }

    setStation(autoState, station) {
      //autoState is boolean
      (this.listType === null ? [station] : station[this.listType].items).forEach((item) => {
        this.fields.forEach((field) => {
          item[field].auto.save(autoState);
          if (station === stationList.activeItem) {
            item[field].auto.refreshInput(false, true);
          }
        });
      });
    }

    setAllStations(autoState) {
      //WARNING: Not required for default course/station, so not implemented
      stationList.items.forEach((station) => { this.setStation(autoState, station); });
    }

    setThisStation(autoState) {
      this.setStation(autoState, stationList.activeItem);
    }

    init() {
      id2element(this, ["autoAllBtn", "autoOneBtn", "manualAllBtn", "manualOneBtn"]);
      this.autoAllBtn.addEventListener("click", () => { this.setAllStations(true); });
      this.autoOneBtn.addEventListener("click", () => { this.setThisStation(true); });
      this.manualAllBtn.addEventListener("click", () => { this.setAllStations(false); });
      this.manualOneBtn.addEventListener("click", () => { this.setThisStation(false); });
    }
  }

  [
    new AutoBtnSet({
      listType: "kites",
      fields: ["kitex", "kitey"],
      autoAllBtn: "autoAllKitePosAll",
      manualAllBtn: "manualAllKitePosAll",
      autoOneBtn: "autoAllKitePosStation",
      manualOneBtn: "manualAllKitePosStation"
    }),
    new AutoBtnSet({
      listType: null,
      fields: ["vpx", "vpy", "heading", "blankMapPage"],
      autoAllBtn: "autoAllHeadingVPAll",
      manualAllBtn: "manualAllHeadingVPAll",
      autoOneBtn: "autoAllHeadingVPStation",
      manualOneBtn: "manualAllHeadingVPStation"
    }),
    new AutoBtnSet({
      listType: "tasks",
      fields: ["solution"],
      autoAllBtn: "autoAllSolutionAll",
      manualAllBtn: "manualAllSolutionAll",
      autoOneBtn: "autoAllSolutionStation",
      manualOneBtn: "manualAllSolutionStation"
    }),
    new AutoBtnSet({
      listType: "tasks",
      fields: ["circlePage", "circlex", "circley", "cdPage", "cdx", "cdy", "cdWidth", "cdHeight", "cdScale"],
      autoAllBtn: "autoAllTaskPositionsAll",
      manualAllBtn: "manualAllTaskPositionsAll",
      autoOneBtn: "autoAllTaskPositionsStation",
      manualOneBtn: "manualAllTaskPositionsStation"
    })
  ].forEach((list) => { list.init(); });

  //Create root level objects and populate with essentials
  //Must do courseList before stationList
  courseList = new CourseList();
  courseList.add();
  stationList = new StationList();
  stationList.add();
  [courseList, stationList].forEach((list) => { list.refreshList(); });

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

  mapsCompiler.setTemplate = function (name) {
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
