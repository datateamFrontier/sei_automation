function getDataFromGoogleSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sheet1");
  const [header, ...data] = sheet.getDataRange().getDisplayValues();
  const choices = {}
  header.forEach(function(title, index){
    choices[title] = data.map(row => row[index]).filter(e => e !== '');
  });
  return choices;
}

function populateGoogleForms(){
  const GOOGLE_FORM_ID = "1hSUiXsMCsnTjLAleT-NFKEllayTMsfs2H_EzH0dCTps";
  const googleForm = FormApp.openById(GOOGLE_FORM_ID);
  const items = googleForm.getItems();
  const choices = getDataFromGoogleSheets();
  items.forEach(function(item){
    const itemTitle = item.getTitle();
    if (itemTitle in choices){
      const itemType = item.getType();
      switch (itemType){
        case FormApp.ItemType.CHECKBOX:
          item.asCheckboxItem().setChoiceValues(choiches[itemTitle]);
          break;
        case FormApp.ItemType.LIST:
          item.asListItem().setChoiceValues(choices[itemTitle]);
          break;
        case FormApp.ItemType.MULTIPLE_CHOICE:
          item.asMultipleChoiceItem.setChoiceValues(choices[itemTitle]);
          break;
        default:
          Logger.log("Ignore question", itemTitle);
      }
    }
  });
  SpreadsheetApp.getActiveSpreadsheet().toast("Google Form Updated");
}
