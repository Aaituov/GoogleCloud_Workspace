// Primary function to work with values in the table.
function extractDataAndCreateNewSheet() {
    var sourceSheetName = 'test'; // Replace with the name of your source sheet
    var targetSheetName = 'test_result'; // Replace with the name for your new sheet
  
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = spreadsheet.getSheetByName(sourceSheetName);
  
    var targetSheet = createOrClearSheet(spreadsheet, targetSheetName);
    var data = getDataFromSheet(sourceSheet);
  
    var headers = ['First name', 'Last name', 'University name', 'Student ID', 'Profile Link', 'Badge number'];
    writeHeadersToSheet(targetSheet, headers);
  
    processAndWriteDataToSheet(data, sourceSheet, targetSheet);
  }
  
  // Function to clear the output sheet
  function createOrClearSheet(spreadsheet, sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
    sheet.clear();
    return sheet;
  }
  
  function getDataFromSheet(sheet) {
    var dataRange = sheet.getDataRange();
    return dataRange.getValues();
  }
  
  function writeHeadersToSheet(sheet, headers) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  function processAndWriteDataToSheet(data, sourceSheet, targetSheet) {
    var headerRow = data[0];
    var columnIndexMap = createColumnIndexMap(headerRow);
  
    for (var i = 1; i < data.length; i++) {
      var rowData = data[i];
      var extractedData = extractDataFromRow(rowData, columnIndexMap);
      writeDataToTargetSheet(targetSheet, extractedData, sourceSheet);
    }
  }
  
  function createColumnIndexMap(headerRow) {
    var columnIndexMap = {};
    for (var i = 0; i < headerRow.length; i++) {
      columnIndexMap[headerRow[i]] = i;
    }
    return columnIndexMap;
  }

  function isProperHttpLink(link) {
  // Define a regular expression pattern for a basic HTTP link
  var httpLinkRegex = /(https?:\/\/\S+)/;

  // Test if the provided link matches the pattern
  var isProperLink = httpLinkRegex.test(link);

  return isProperLink;
  }
  
  function extractDataFromRow(rowData, columnIndexMap) {
    var timestampIndex = columnIndexMap['Timestamp'];
    var emailIndex = columnIndexMap['Your email'];
    var nameIndex = columnIndexMap['Your FIRST NAME and LAST NAME'];
    var profileLinkIndex = columnIndexMap['Copy paste Link to your Skills badge at CloudSkills boost profile. For example here https://www.cloudskillsboost.google/public_profiles/4b57394c-8f9a-4bea-9bf6-f15e3abdbd18'];
    var universityNameIndex = columnIndexMap['Your University  name'];
    var studentIDIndex = columnIndexMap['Your student ID'];
  
    var timestamp = rowData[timestampIndex];
    var email = rowData[emailIndex];
    var name = rowData[nameIndex];
    var profileLink = rowData[profileLinkIndex];
    var universityName = rowData[universityNameIndex];
    var studentID = rowData[studentIDIndex];
    var names = name.split(' ');
    var firstName = names[0];
    var lastName = names.slice(1).join(' ');
    if (!isProperHttpLink(profileLink)) {
      Logger.log('Invalid profile link - unable to request: ' + profileLink);
      return [firstName, lastName, universityName, studentID, profileLink, 'Unable to request']; // Or handle it in a way that suits your needs
    }

    if (hasMultipleUrls(profileLink)) {
      profileLink = transformProfileUrl(extractFirstUrl(profileLink));
    } else {
      profileLink = transformProfileUrl(profileLink);
    }
    
    // Check if the profile link is in the correct format
    if (isValidProfileLink(profileLink)) {
      var badgeNumber = extractBadgeNumberFromHTML(profileLink);
  
      return [firstName, lastName, universityName, studentID, profileLink, badgeNumber];
    } else {
      // Log or handle the case of an invalid profile link
      Logger.log('Invalid profile link: ' + profileLink);
      return [firstName, lastName, universityName, studentID, profileLink, 'Invalid profile link']; // Or handle it in a way that suits your needs
    }
  }
  
  function extractFirstUrl(inputString) {
    // Use a regular expression to extract the first URL
    var urlRegex = /https:\/\/www\.cloudskillsboost\.google\/public_profiles\/[a-f\d-]+/g;
    var match = inputString.match(urlRegex);
  
    // Return the first URL or null if none is found
    return match ? match[0] : null;
  }
  
  function isValidProfileLink(profileLink) {
    // Define a regular expression to match valid profile links
    var validProfileLinkRegex = /https:\/\/(www\.cloudskillsboost\.google\/public_profiles\/[a-f\d-]+)/;
  
    // Test if the provided profile link matches the valid format
    return validProfileLinkRegex.test(profileLink);
  }
  
  function hasMultipleUrls(inputString) {
    // Use a regular expression to check if the input contains multiple URLs
    var urlRegex = /https:\/\/www\.cloudskillsboost\.google\/public_profiles\/[a-f\d-]+\/badges\/\d+/g;
    return inputString.match(urlRegex) !== null;
  }
  
  function transformProfileUrl(inputUrl) {
    return inputUrl.replace(/\/badges\/\d+$/, '');
  }
  
  function extractBadgeNumberFromHTML(profileLink) {
    try {
      var regex = /(<div\s+class=['"]profile-badge['"]>|<div\s+class=['"]badge['"]>)/gi; // Update this according to your HTML structure
      var htmlContent = UrlFetchApp.fetch(profileLink.trim()).getContentText();
      var match = htmlContent.match(regex);
      return match ? match.length : 'N/A';
    } catch (e) {
      // Log an error if there is an HTTP error
      if (e.toString().indexOf('Exception: ') === 0) {
        var errorDetails = e.toString().substring('Exception: '.length);
        Logger.log('HTTP error for profile link ' + profileLink + ': ' + errorDetails);
      } else {
        // Log other types of errors
        Logger.log('Error for profile link ' + profileLink + ': ' + e.toString());
      }
      return 'Error';
    }
  }
  
  function writeDataToTargetSheet(targetSheet, rowData, sourceSheet) {
    targetSheet.appendRow(rowData);
  }
  