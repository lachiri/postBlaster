var scriptTitle = "postBlaster Script V1.0 (10/24/13)";
var scriptName = 'postBlaster';
var scriptTrackingId = 'UA-45137515-1';
var waitingIconId = '0B7-FEGXAo-DGT0g5THB3T0hpTlU';
var waitingImageUrl = 'https://drive.google.com/uc?export=download&id='+this.waitingIconId;

function onInstall() {
  onOpen();
}

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var menuEntries = [];
  var properties = ScriptProperties.getProperties()
  menuEntries.push({name: "What is postBlaster?", functionName: "postBlaster_whatIs"});
  menuEntries.push(null);
  if (!properties.initialized) {
    menuEntries.push({name: "Complete Initial Installation", functionName: "postBlaster_onInstall"});
  }
  if (properties.initialized) {
    menuEntries.push({name: "Configure Settings", functionName: "postBlaster_configureSettings"});
    menuEntries.push({name: "Bypass Time Trigger and Manually Email New Posts", functionName: "postBlaster_checkAndSend"});
  }
  ss.addMenu("postBlaster", menuEntries);
}

function postBlaster_setTimeTrigger(){
  var ss = SpreadsheetApp.getActive();
  var triggers = ScriptApp.getProjectTriggers()
  var found = false;
  for (var i=0; i<triggers.length; i++) {
    if (triggers[i].getHandlerFunction()=="postBlaster_checkAndSend"){
      found = true;
      break;
    } 
  }
  if (found == false){
    ScriptApp.newTrigger("postBlaster_checkAndSend").timeBased().everyMinutes(5).create()
  }
  var username = Session.getActiveUser().getEmail()
  ScriptProperties.setProperty('timeTrigger', username)
  return username;
}


function postBlaster_onInstall(){
  setSid();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var emailsSheet = postBlaster_getSheetById(properties.emailsSheetId, ss);
  if ((!properties.emailsSheetId)||(!emailsSheet)){
    try {
      var emailsSheet = ss.insertSheet('Email Distribution List');
      var emailsColumnNames = ["Emails (Required)", "Name (Optional)", "Additional Header 1 (Optional)"];
      var emailsHeadersRange = emailsSheet.getRange(1, 1, 1, emailsColumnNames.length);
      emailsHeadersRange.setValues([emailsColumnNames]).setBackground("#d8d8d8").setFontWeight("bold");
      var emailsSampleRange = emailsSheet.getRange(2, 1, 3, emailsColumnNames.length)
      var emailsSampleData = [["email1@sampledomain.org", "Danielle Phillips", ""],
                              ["email2@sampledomain.org", "Joe Smith", ""],
                              ["group1@sampledomain.org", "Staff Distribution Group", ""]]
      emailsSampleRange.setValues(emailsSampleData);
    } catch(err) {
      // leave for now ... 
    }
    var logSheet = postBlaster_getSheetById(properties.logSheetId, ss);
    if ((!properties.logSheetId)||(!logSheet)){
       logSheet = insertLogSheet(ss);
    } 
    var emailsSheetId = emailsSheet.getSheetId();
    ScriptProperties.setProperty('emailsSheetId', emailsSheetId);
    ScriptProperties.setProperty('emailsColumn', "Email");
    var now = Number(new Date ()).toString();
    ScriptProperties.setProperty('initTime', now);
    ScriptProperties.setProperty('initialized', 'true');
  }
  onOpen();
}


function insertLogSheet(ss) {
  var logSheet = ss.insertSheet('Sent Emails');
  var logSheetColumnNames = ["Date", "URL","Title", "Author", "Key", "Status"];
  var logSheetHeadersRange = logSheet.getRange(1, 1, 1, logSheetColumnNames.length);
  logSheetHeadersRange.setValues([logSheetColumnNames]).setBackground("#d8d8d8").setFontWeight("bold");
  var logSheetId = logSheet.getSheetId();
  ScriptProperties.setProperty('logSheetId', logSheetId);
  return logSheet;
}

function postBlaster_validateUrl(e) {
  var app = UiApp.getActiveApplication();
  var announcePageWarning = app.getElementById('badPageWarning');
  var url = e.parameter.announcePageUrl;
  if ((!url)||(url == "")) {
    announcePageWarning.setVisible(false);
    return app;
  }
  try {
    var loadedPage = SitesApp.getPageByUrl(url);
    var type = loadedPage.getPageType().toString();
    if (type == "AnnouncementsPage"){
      announcePageWarning.setVisible(false);
    } else {
      announcePageWarning.setVisible(true);
    }
  } catch(err) {
    announcePageWarning.setVisible(true);
  }
  
  return app;
}

function postBlaster_configureSettings(){
  var properties = ScriptProperties.getProperties()
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = postBlaster_getSheetById(properties.logSheetId, ss);
  if (!logSheet) {
    insertLogSheet(ss)
  }
  var app = UiApp.createApplication().setTitle("Configure postBlaster Settings").setHeight(400);
  
  var waitingIcon = app.createImage(waitingImageUrl)
  .setHeight("150px")
  .setWidth("150px")
  .setStyleAttribute('position', 'absolute')
  .setStyleAttribute('left', '35%')
  .setStyleAttribute('top', '35%')
  .setVisible(false);
  
  var outerScrollPanel = app.createScrollPanel().setHeight("380px");
  var panel = app.createVerticalPanel();
  var grid = app.createGrid(7,2).setCellPadding(8).setStyleAttribute('backgroundColor', '#D1D0CE');
  
  grid.setWidget(0, 0, app.createLabel('Copy and paste the URL of your Announcements Page.'));
  
  var announcePageUrl = properties.announcePageUrl;
  var announcePageMiniPanel = app.createVerticalPanel();
  var announcePageTextBox = app.createTextArea().setName('announcePageUrl').setId('announcePageUrl').setWidth("250px");
  var announcePageWarning = app.createLabel("Invalid or inaccessible Google Sites Announcement Page URL").setStyleAttribute('fontSize', '11px').setStyleAttribute('marginTop', '4px').setId('badPageWarning').setStyleAttribute('color', 'red').setVisible(false);
  var urlHandler = app.createServerHandler('postBlaster_validateUrl').addCallbackElement(announcePageTextBox);
  announcePageTextBox.addKeyUpHandler(urlHandler);
  announcePageMiniPanel.add(announcePageTextBox).add(announcePageWarning);
  if (properties.announcePageUrl){
    announcePageTextBox.setValue(announcePageUrl); 
  }
  
  grid.setWidget(0, 1, announcePageMiniPanel);
  
  grid.setWidget(1, 0, app.createLabel('Select the Sheet that contains your Email Distribution List:'));
  var sheets = ss.getSheets();
  var sheetOptions = app.createListBox().setId('sheetChooser').setName('emailsSheetId');
  var allSheetIds = []
  for (var i=0; i<sheets.length; i++) {
    if (sheets[i].getSheetId()!=properties.logSheetId){
      allSheetIds.push(sheets[i].getSheetId());
      sheetOptions.addItem(sheets[i].getSheetName(), sheets[i].getSheetId()); 
    }
  }
  if (properties.emailsSheetId){
    var index = allSheetIds.indexOf(parseInt(properties.emailsSheetId));
    sheetOptions.setSelectedIndex(index);
  }
  
  var handler2 = app.createServerHandler('postBlaster_refreshHeaderOptions').addCallbackElement(panel);
  sheetOptions.addChangeHandler(handler2);
  grid.setWidget(1, 1, sheetOptions)
  
  var button3 = app.createButton('Select and Close');
  var handler3 = app.createServerClickHandler('postBlaster_saveSettings').addCallbackElement(panel);
  var clientHandler = app.createClientHandler().forTargets(panel).setStyleAttribute('opacity', '0.5').forTargets(waitingIcon).setVisible(true).forTargets(button3).setEnabled(false);
  button3.addClickHandler(handler3).addClickHandler(clientHandler);
  
  grid.setWidget(2, 0, app.createLabel('Select the Column that contains your Email Distribution List:'));
  var columnOptions = app.createListBox().setId('columnChooser').setName('emailsColumn');
  postBlaster_refreshHeaderOptions();
  grid.setWidget(2, 1, columnOptions);
  
  grid.setWidget(3, 0, app.createLabel('Use these tokens when building your custom subject line and footer:'));
  var tokenOptions = app.createLabel('$title  -  $author  -  $date  -  $announcementUrl  -  $siteUrl');
  grid.setWidget(3, 1, tokenOptions);
  
  grid.setWidget(4, 0, app.createLabel('Use this handy html tag to hyperlink the url tokens:'));
  var handyHtml = app.createLabel('<a href= "URL">Link Text</a>');
  grid.setWidget(4, 1, handyHtml);
  
  var subjectString = properties.subjectString;
  grid.setWidget(5, 0, app.createLabel('Build your custom subject line here:'));
  var subjectStringTextBox = app.createTextArea().setName('subjectString').setId('subjectString').setWidth("250px");
  if (properties.subjectString){
    subjectStringTextBox.setValue(subjectString); 
  }
  grid.setWidget(5, 1, subjectStringTextBox);

  var footerString = properties.footerString;
  grid.setWidget(6, 0, app.createLabel('Build your custom footer here:'));
  var footerStringTextBox = app.createTextArea().setName('footerString').setId('footerString').setWidth("250px");
  if (properties.footerString){
    footerStringTextBox.setValue(footerString); 
  }
  grid.setWidget(6, 1, footerStringTextBox);
  
  panel.add(grid);
  panel.add(button3);
  outerScrollPanel.add(panel);
  app.add(outerScrollPanel);
  app.add(waitingIcon);
  ss.show(app);
  
}
 
function postBlaster_refreshHeaderOptions(e) {
  var properties = ScriptProperties.getProperties()
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.getActiveApplication();
  if (!e){
    var value = parseInt(properties.emailsSheetId);
    if (!value){
      value = ss.getSheets()[0].getSheetId();
    }
  } else {
    var value = parseInt(e.parameter.emailsSheetId);
  }
  var sheet = postBlaster_getSheetById(value, ss);
  var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  var headers = headersRange.getValues()[0];
  var columnOptions = app.getElementById('columnChooser');
  columnOptions.clear();
  
  for (var i=0; i<headers.length; i++) {
    columnOptions.addItem(headers[i]);  
  }
  if (properties.emailsColumn){
    var index = headers.indexOf(properties.emailsColumn);
    columnOptions.setSelectedIndex(index);
  }
  
  return app
}

function postBlaster_saveSettings(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  ScriptProperties.setProperty('ssId', ssId);
  var app = UiApp.getActiveApplication();
  var announcePageUrl = e.parameter.announcePageUrl;
  ScriptProperties.setProperty('announcePageUrl', announcePageUrl)
  var emailsSheetId = e.parameter.emailsSheetId
  ScriptProperties.setProperty('emailsSheetId', emailsSheetId)
  var emailsColumn = e.parameter.emailsColumn;
  ScriptProperties.setProperty('emailsColumn', emailsColumn)
  var subjectString = e.parameter.subjectString;
  ScriptProperties.setProperty('subjectString', subjectString)
  var footerString = e.parameter.footerString;
  ScriptProperties.setProperty('footerString', footerString)
  
  var timeTrigger = ScriptProperties.getProperty('timeTrigger')
  if (!timeTrigger){
    var timeTrigger = postBlaster_setTimeTrigger()
  }
  Browser.msgBox("This script will run as the person who originally installed it.  Every 5 minutes postBlaster will find new posts and send emails from the following address: " + timeTrigger + ".")

  app.close()
  return app
 
}

function postBlaster_getSubPageObjects(url){
  var properties = ScriptProperties.getProperties()
  var announcePageUrl = properties.announcePageUrl
  var announcements = SitesApp.getPageByUrl(announcePageUrl).getAnnouncements();
  var objects = [];
  
  for (var i = 0; i<announcements.length; i++){
    objects[i] = new Object();
    objects[i].date = announcements[i].getDatePublished();
    objects[i].url = announcements[i].getUrl();
    objects[i].title = announcements[i].getTitle();
    objects[i].author = announcements[i].getAuthors().join(',');
    objects[i].key = Number(announcements[i].getDatePublished()) 
    objects[i].html = announcements[i].getHtmlContent();
  }
  return objects
}

function postBlaster_replaceTokens(subPageObject, properties, siteUrl) {
  var ss = SpreadsheetApp.openById(properties.ssId);
  var timeZone = ss.getSpreadsheetTimeZone();
  var titleToken = new RegExp("\\$title", "g");
  var authorToken = new RegExp("\\$author", "g");
  var dateToken = new RegExp("\\$date", "g");
  var announcementUrlToken = new RegExp("\\$announcementUrl", "g");
  var siteUrlToken = new RegExp("\\$siteUrl", "g");
  var datePublishedAsString = Utilities.formatDate(subPageObject.date, timeZone, "M/dd/yyyy H:m a");
  
  var subjectString = properties.subjectString;
  subjectString = subjectString.replace(titleToken, subPageObject.title);
  subjectString = subjectString.replace(authorToken, subPageObject.author);
  subjectString = subjectString.replace(dateToken, datePublishedAsString);
  subjectString = subjectString.replace(announcementUrlToken, subPageObject.url);
  subjectString = subjectString.replace(siteUrlToken, siteUrl);
  
  var footerString = properties.footerString;
  footerString = footerString.replace(titleToken, subPageObject.title);
  footerString = footerString.replace(authorToken, subPageObject.author);
  footerString = footerString.replace(dateToken, datePublishedAsString);
  footerString = footerString.replace(announcementUrlToken, subPageObject.url);
  footerString = footerString.replace(siteUrlToken, siteUrl);
  
  var returnObj = {};
  returnObj.subject = subjectString;
  returnObj.footer = footerString;
  return returnObj;
}

function postBlaster_getSentPageKeys()  {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var logSheetId = properties.logSheetId
  var logSheet = postBlaster_getSheetById(logSheetId, ss);
  var data = []
  if (logSheet.getLastRow()>1){
    var range = logSheet.getRange(2, 1, logSheet.getLastRow()-1, logSheet.getLastColumn());
    data = NVSL.getRowsData(logSheet, range);
  }
  var keys = []
  
  for (var i = 0; i<data.length; i++){
    keys.push(data[i].key);
  }
  return keys
}

function postBlaster_getSheetById(sheetId, ss) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0; i<sheets.length; i++){
    if (sheets[i].getSheetId()==sheetId) {
      return sheets[i];
    }
  }
  return;
}

function postBlaster_getRecipients(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties()
  var emailsSheetId = properties.emailsSheetId
  var emailsSheet = postBlaster_getSheetById(emailsSheetId, ss);
  var emailsColumn = properties.emailsColumn;
  var normalizedEmailHeader = NVSL.normalizeHeaders([emailsColumn])[0];
  var emailsSheetHeaders = emailsSheet.getRange(1, 1, 1, emailsSheet.getLastColumn()).getValues()[0];
  var emailsDataRange = emailsSheet.getRange(2, 1, emailsSheet.getLastRow()-1, emailsSheet.getLastColumn());
  var emails = []
  var emailsData = NVSL.getRowsData(emailsSheet, emailsDataRange); 
  for (var i = 0; i<emailsData.length; i++){
    emails.push(emailsData[i][normalizedEmailHeader]);
  }
  return emails
}

function postBlaster_writeBackToSheet(stuff){
  if (stuff.length>0){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var properties = ScriptProperties.getProperties();
    var announcePageUrl = properties.announcePageUrl;
    var subPageObjects = postBlaster_getSubPageObjects(announcePageUrl);
    var logSheetId = properties.logSheetId;
    var logSheet = postBlaster_getSheetById(logSheetId, ss);
    var headersRange = logSheet.getRange(1, 1, 1, logSheet.getLastColumn());
    NVSL.setRowsData(logSheet, stuff, headersRange, logSheet.getLastRow()+1);
  }
}

function postBlaster_checkAndSend(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var initTime = parseInt(properties.initTime);
  var announcePageUrl = properties.announcePageUrl;
  var siteUrl = SitesApp.getSiteByUrl(announcePageUrl).getUrl()
  var subPageObjects = postBlaster_getSubPageObjects(announcePageUrl);
  var logSheetId = properties.logSheetId;
  var logSheet = postBlaster_getSheetById(logSheetId, ss);
  if (!logSheet) {
    logSheet = insertLogSheet(ss)
  }
  var alreadySentKeys = postBlaster_getSentPageKeys();
  var emails = postBlaster_getRecipients();
  var stuff = [];
  
  for (var i = 0; i<subPageObjects.length; i++){
    if ((alreadySentKeys.indexOf(subPageObjects[i].key)==-1)&&(subPageObjects[i].key>initTime)) {
      try {
        postBlaster_logPostEmailed();
      } catch(err) {
      }
      var returnObj = postBlaster_replaceTokens(subPageObjects[i], properties, siteUrl);
      for (var j = 0; j<emails.length; j++){
        try{
          MailApp.sendEmail(emails[j], returnObj.subject,'',{htmlBody:subPageObjects[i].html+'<p><p>'+returnObj.footer});
          postBlaster_logEmailSent();
        }catch(err){
        }
      }
      var thisObj = new Object;
      for (var key in subPageObjects[i]){
        if (key!='html'){
          thisObj[key] = subPageObjects[i][key];
        }
      }
      var timestamp = new Date ();
      thisObj.status = 'sent on '+ timestamp;
      stuff.push(thisObj);
    }
  }
  postBlaster_writeBackToSheet(stuff)
}

