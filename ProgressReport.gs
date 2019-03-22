/* PROGRESS REPORT GENERATION SCRIPT
 * AUTHOR: Agatha Downey
 * LAST UPDATED: 21 March 2019
 */

// This function triggers on submission of the form linked to this spreadsheet
function onFormSubmit(e) {
  //---------------------------------------------------------------------------------------------
  
  // Opens the attached spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  // Gets the data from the spreadsheet
  var data = sheet.getDataRange().getValues();
  // Takes the first data row as the headers (names)
  var headers = data[0];
  // Gets the number of the row that was just modified and triggered this script
  var newest_row_index = e.range.getRow();
  // Opens the data from trigger-row
  var newest_row = data[newest_row_index - 1];
  
  //---------------------------------------------------------------------------------------------
  
  // Finds the column that holds the "Name" information for the student (used to name the file)
  var name_column = -1
  for(var i = 0; i < headers.length; i++){
    if(headers[i] == 'Name'){
      name_column = i
    }
  }
  var name = newest_row[name_column]
  
  // Gets the current year
  var now = new Date();
  var year = now.getFullYear();
  
  var docName = name + ' Spr ' + year + ' Progress Report'
  
  //---------------------------------------------------------------------------------------------
  /* This section deletes all previous files under the name of the student who just triggered the script;
   * This is necessary because if a student submits their response, and then updates their answers, two
   * versions of the progress report will be created, and it will be unclear which one is correct. This
   * section ensures only the most recent data from the student will be used in the progress report
   */
  
  var thisFileId = sheet.getParent().getId();
  var thisFile = DriveApp.getFileById(thisFileId);
  var parentFolder = thisFile.getParents().next();
  var allFiles = parentFolder.getFiles();
  var toDelete = []

  while (allFiles.hasNext()) {
    var iterFile = allFiles.next();
    if (iterFile.getName() == docName){
      var deleteFile = iterFile.getId();
      Logger.log(deleteFile);
      toDelete.push(deleteFile);
    }
  };
  
  for(var i = 0; i < toDelete.length; i++){
    Logger.log(toDelete[i])
    var result = Drive.Files.remove(toDelete[i])
    Logger.log(result);
  }
  
  //---------------------------------------------------------------------------------------------
  // This is the part of the script that creates the progress report
  
  /* Opens the template file
   * IMPORTANT: THIS ID MUST CORRESPOND TO THE EXACT FILE BEING USED AS THE TEMPLATE
   * THE ID CAN BE FOUND IN THE URL WHEN YOU OPEN THE TEMPLATE DOCUMENT
   */
  var templateId = 'REPLACE_THIS_WITH_TEMPLATE_DOC_ID';
  
  // Maks a copy of the template file
  var documentId = DriveApp.getFileById(templateId).makeCopy().getId();
  
  // Names the newly created file with the name of the student who submitted the information and the year
  DriveApp.getFileById(documentId).setName(docName);
  
  // Gets the document body as a variable
  var body = DocumentApp.openById(documentId).getBody();
  
  for(var i = 0; i < headers.length; i++){
      body.replaceText('##' + headers[i] + '##', newest_row[i])
  }
}
