function apiEndPointGenerator() {
     const userData = getUserData();
   
     const userQuery = userData.keywords;
     const userDateFrom = userData.dateFrom;
     const userDateTo = userData.dateTo;
     const userLang = userData.language;
     const userSortBy = userData.sortBy;
   
     const API_KEY = `YOUR_API_GOES_HERE`;
     const url = `https://newsapi.org/v2/everything?q=${userQuery}&from=${userDateFrom}&to=${userDateTo}&language=${userLang}&sortBy=${userSortBy}&apiKey=${API_KEY}`;
   
     return url;
   }
   
   function fetchNews() {
     const endPoint = apiEndPointGenerator();
     let response = UrlFetchApp.fetch(endPoint, {
       method: "GET",
     });
     let data = response.getContentText();
     let jsonData = JSON.parse(data);
     let reqData = jsonData.articles;
   
     return reqData.map((article) => {
       return {
         title: article.title,
         description: article.description,
         url: article.url,
         content: article.content,
         author: article.author,
       };
     });
   }
   
   function createLayout() {
     let currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
   
     currentSheet.getRange("A1").setValue("Keywords");
     currentSheet.getRange("B1").setValue("NewsFrom");
     currentSheet.getRange("C1").setValue("NewsTo");
     currentSheet.getRange("D1").setValue("Language");
     currentSheet.getRange("E1").setValue("SortBy");
     
     setDatesAndDropdowns();
   
     currentSheet.getRange("A4").setValue("Title");
     currentSheet.getRange("B4").setValue("Description");
     currentSheet.getRange("C4").setValue("URL");
     currentSheet.getRange("D4").setValue("Content");
     currentSheet.getRange("E4").setValue("Author");
   
   }
   
   function clearData() {
     const sheetName = "Sheet1";
     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
     if (!sheet) {
       throw new Error(`Sheet ${sheetName} not found.`);
     }
   
     const dataRange = sheet.getDataRange();
     const numRows = dataRange.getNumRows();
     const numCols = dataRange.getNumColumns();
   
     if (numRows > 4 || numCols > 5) {
       // Define range to clear all data except A1 to E4
       const range = sheet.getRange(5, 1, numRows - 4, numCols);
       range.clearContent();
       console.log(`Cleared ${range.getNumRows()} rows in ${sheetName}.`);
     } else {
       console.log(`No data to clear in ${sheetName}.`);
     }
   }
   
   function createSheet1() {
     let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
     
     // If Sheet1 does not exist, create it
     if (sheet1 == null) {
       sheet1 = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
       sheet1.setName("Sheet1");
       return true;
     }
     
     // If Sheet1 already exists, return false
     return false;
   }
   
   // function to grab data from user
   function getUserData() {
     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
     const data = sheet.getRange(2, 1, 1, 5).getValues()[0];
     const [newsKeywords, newsFrom, newsTo, newsLanguage, newsSortBy] = data;
   
     // Setting the time zone to "+0545"
     const timeZone = "GMT+5:45";
     const dateFormat = "yyyy-MM-dd";
   
     // Formatting the 'newsFromDate' and 'newsToDate' in 'yyyy-mm-dd' format
     const formattedNewsFromDate = Utilities.formatDate(
       new Date(newsFrom),
       timeZone,
       dateFormat
     );
     const formattedNewsToDate = Utilities.formatDate(
       new Date(newsTo),
       timeZone,
       dateFormat
     );
   
     const userData = {
       keywords: newsKeywords,
       dateFrom: formattedNewsFromDate,
       dateTo: formattedNewsToDate,
       language: newsLanguage,
       sortBy: newsSortBy,
     };
   
     return userData;
   }
   
   function writeNewsToSheet() {
     const data = fetchNews();
     const sheetName = "Sheet1";
     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
     if (!sheet) {
       throw new Error(`Sheet ${sheetName} not found.`);
     }
   
     const startRow = 5;
     const startColumn = 1;
     const numRows = data.length;
     const numColumns = Object.keys(data[0]).length;
   
     const range = sheet.getRange(startRow, startColumn, numRows, numColumns);
     range.setValues(data.map(article => Object.values(article)));
   
     console.log(`Wrote ${numRows} rows to ${sheetName} starting from ${startRow}, ${startColumn}.`);
   }
   
   function setDatesAndDropdowns() {
     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
     
     // Set one week earlier date in B2 cell
     const oneWeekAgo = new Date();
     oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
     sheet.getRange("B2").setValue(oneWeekAgo);
     
     // Set today's date in C2 cell
     const today = new Date();
     sheet.getRange("C2").setValue(today);
     
     // Set data validation for B2 and C2 cells
     const dateValidation = SpreadsheetApp.newDataValidation().requireDate().build();
     sheet.getRange("B2:C2").setDataValidation(dateValidation);
     
     // Set dropdown menus in D2 and E2 cells
     const languageDropdown = SpreadsheetApp.newDataValidation().requireValueInList(["en", "ar"], true).build();
     const sortOptionsDropdown = SpreadsheetApp.newDataValidation().requireValueInList(["relevancy", "popularity", "publishedAt"], true).build();
     sheet.getRange("D2").setDataValidation(languageDropdown);
     sheet.getRange("E2").setDataValidation(sortOptionsDropdown);
     
     // Set default values in D2 and E2 cells
     sheet.getRange("D2").setValue("en");
     sheet.getRange("E2").setValue("relevancy");
   }
   
   function setFont() {
     const sheetName = "Sheet1";
     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
     if (!sheet) {
       throw new Error(`Sheet ${sheetName} not found.`);
     }
   
     const font = "Verdana";
     const fontSize = 12;
   
     const range = sheet.getDataRange().setFontFamily(font).setFontSize(fontSize);
     range.getCell(1, 1).offset(0, 0, 1, sheet.getLastColumn()).setFontWeight("bold");
     range.getCell(4, 1).offset(0, 0, 1, sheet.getLastColumn()).setFontWeight("bold");
   
     console.log(`Font set to ${font}, font size set to ${fontSize}, and rows 1 and 4 set to bold in ${sheetName}.`);
   }
   
   function renameSheetFile() {
     const newName = "AB NewsAPI x Google Sheets";
     const sheet = SpreadsheetApp.getActiveSpreadsheet();
     const fileId = sheet.getId();
     const file = DriveApp.getFileById(fileId);
     file.setName(newName);
     console.log(`Renamed file to "${newName}".`);
   }
   
   function workFlow() {
     renameSheetFile();
     clearData();
     writeNewsToSheet();
     setFont();
   }
   
   function onOpen(){
     createSheet1();
     createLayout();
     let ui = SpreadsheetApp.getUi();
     let menu = ui.createMenu('Fetch Menu');
     menu.addItem('Fetch News', 'workFlow').addToUi();
   }