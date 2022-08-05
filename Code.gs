var ui = SpreadsheetApp.getUi();
var destSheet;
var rule;
var newCriteriaValue;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Re-use Conditonal Formatting')
    .addItem('Open the sidebar options', 'showSidebar')
    .addItem('About', 'about')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
    .setTitle('Re-use Conditonal Formatting');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function getSheetNames() {
  var names = SpreadsheetApp.getActive().getSheets();
  return names.map(x => { return x.getName() });
}

function getSelectedSheet(selection) {
  try {
    var selectedSheet = SpreadsheetApp.getActive().getSheetByName(selection);
    storeSelectedSheet(selection);
    SpreadsheetApp.setActiveSheet(selectedSheet);
  } catch (e) {
    console.log(e);
  }
}

function checkRangeSelection() {
  try {
    rule = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getConditionalFormatRules()[0];
    storeOriginalSheet(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName());
    var ranges = rule.getRanges();
    var selectedRange = SpreadsheetApp.getActiveSheet().getSelection().getActiveRange().getA1Notation();

    for (var i = 0; i < ranges.length; i++) {

      if (selectedRange == ranges[i].getA1Notation()) {
        var criteriaValue = rule.getBooleanCondition().getCriteriaValues().length < 1 ? 'EMPTY' : rule.getBooleanCondition().getCriteriaValues();
        var response = ui.alert("[Selected Range]\n" + selectedRange +
          "\n\nThis range has these details:\n" +
          "\n[Type]\n" + rule.getBooleanCondition().getCriteriaType() +
          "\n\n[Value]\n" + criteriaValue + '\n\nDo you want to update it and/or apply it to a new sheet range?',
          ui.ButtonSet.YES_NO);

        if (response == ui.Button.YES) {
          var valueNew = ui.prompt("[Current Criteria Value]\n" + criteriaValue +
            "\n\nChange it/press the \"Ok\" button to use the default criteria value:\n");
          valueNew.getResponseText() == '' ? storeNewCriteriaValue(rule.getBooleanCondition().getCriteriaValues().toString()) : storeNewCriteriaValue(valueNew.getResponseText());
        } else {
          ui.alert('Cancelled');
        }

      } else {
        ui.alert('Selected range \"' + selectedRange + '\" doesn\'t contain any criteria.');
      }
    }

  } catch {
    ui.alert('Selected sheet doesn\'t contain any Conditonal Formatting');
  }
}

function applyConditionalFormatting() {
  try {
    passNewCriteriaValue();
    var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(passDestionationSheet()); //pass stored destination sheet to the global variable rule
    var range = destinationSheet.getRange(destinationSheet.getActiveRange().getA1Notation());
    rule = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(passOGSheet()).getConditionalFormatRules()[0];
    var copiedrule = rule.copy().withCriteria(rule.getBooleanCondition().getCriteriaType(), [passNewCriteriaValue()]).setRanges([range]).build();
    var rules = destinationSheet.getConditionalFormatRules();
    rules.push(copiedrule);
    destinationSheet.setConditionalFormatRules(rules);
  } catch (e) {
    ui.alert("Oppps! Possible issues:\n\n1. Skipped the \"Step 1\"\n2. No \"Destination Sheet\" selected\n\nPlease follow the steps.")
    console.log(e)
  }
  reset();
  showSidebar();
}

function storeSelectedSheet(destSheet) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('destSheet', destSheet);
  } catch (err) {
    Logger.log('Failed with error %s', err.message);
  }
}

function storeOriginalSheet(originSheet) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('originSheet', originSheet);
  } catch (err) {
    Logger.log('Failed with error %s', err.message);
  }
}

function passDestionationSheet() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const destination = scriptProperties.getProperty('destSheet');
    return destination;
  } catch (err) {
    Logger.log('Failed with error %s', err.message);
  }
}

function passOGSheet() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const ogsheet = scriptProperties.getProperty('originSheet');
    return ogsheet;
  } catch (err) {
    Logger.log('Failed with error %s', err.message);
  }
}

function storeNewCriteriaValue(newCriteriaValue) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('newCriteria', newCriteriaValue);
  } catch (err) {
    Logger.log('Failed with error %s', err.message);
  }
}

function passNewCriteriaValue() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const newCriteriaValue = scriptProperties.getProperty('newCriteria');
    return newCriteriaValue;
  } catch (err) {
    Logger.log('Failed with error %s', err.message);
  }
}

function reset() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteAllProperties();
  } catch (err) {
    Logger.log('Failed with error %s', err.message);
  }
}

function about() {
  ui.alert("This is a simple Google Sheets add-on for extracting/editing any conditional formatting & then applying it to another range from existing sheet tab or to another sheet tab in a Google Spreadsheet file.\n\n\nDeveloped by Irvin Jay Guinto • 2022");
}