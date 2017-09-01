
/*-------------------------------------
/ When the SpreadSheet (Contacts) Opens,
/  Make a new menu item called
/  'Gmail Mail Merge' 
/  Add a submenu item: 'Start mail merge 
/   utility'
---------------------------------------*/

function onOpen () {
  var ui = SpreadsheetApp.getUi()
  ui.createMenu('Gmail Mail Merge')
    .addItem('Start mail merge utility', 'main')
    .addToUi()
}

/*-------------------------------------
/ Get all of the Column Headings then
/  populate and return a map of their
/  numerical references (e.g. Email is 2).  
/ 
---------------------------------------*/

function getColumnHeadings (sheet) {
  var headerRange = sheet.getRange(1, 1, 1, sheet.getMaxColumns())
  var rawHeader = headerRange.getValues()
  var header = rawHeader[0]
  var columns = {}

  for (var i in header) {
    var ref = header[i]
    if (header[i] !== '') {
      columns[ref] = i
    }
  }
  return columns
}

/*-------------------------------------
/ Query gmail for the latest draft
/  in the users email. 
/
---------------------------------------*/

function getLatestDraft () {
  var drafts = GmailApp.getDraftMessages()
  if (drafts.length === 0) return null
  else return drafts[0]
}

/*-------------------------------------
/ Main Body of the script.
/  1. Check quota and disclose to user
/   (fail if quota is already met)
/  2. Find the latest draft
/   (confirm it is the correct draft)
/  3. For each contact, replace place
/   holders with specified name/text. 
/
---------------------------------------*/

function main () {
  var quota = MailApp.getRemainingDailyQuota()
  var ui = SpreadsheetApp.getUi()

  ui.alert('Your remaining daily email quota: ' + quota)

  if (quota === 0) {
    ui.alert('You can not send more emails')
    return
  }

  var labelRegex = /{{[\w\s\d]+}}/g
  var sheet = SpreadsheetApp.getActiveSheet()

  var draft = getLatestDraft()
  
  if (draft === null) {
    ui.alert('No email draft found in your Google account. First draft an email.')
    return
  }
  
  var subject = draft.getSubject()
  var htmlBodyRaw = draft.getBody()
  var plainBodyRaw = draft.getPlainBody()
  var emailPlaceHolder = 0
  

  var response = ui.alert('Do you want to send drafted message titled "' + subject + '" to ' + (sheet.getLastRow() - 1) + ' people?', ui.ButtonSet.YES_NO)

  if (response === ui.Button.NO) {
    ui.alert('Stopped')
    return
  }

  var count = 0
  var emailMap = {}

  var columns = getColumnHeadings(sheet)
  
  if(columns['Email'] != null){
   emailPlaceHolder = columns['Email'] 
  }
  else{
   emailPlaceHolder = 2
  }
  
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn())
  var data = dataRange.getValues()

  for (var i in data) {
    if(data[i] != null){
     var row = data[i]
     var email = row[emailPlaceHolder]
 
     if (email !== null && !emailMap[email] && email !== '' && email !== ' ' && email !== "undefined") {
       var htmlBody = htmlBodyRaw.replace(labelRegex, function (k) {
         var label = k.substring(2, k.length - 2)
         Logger.log('Replaced ' + k + ' with ' + row[columns[label]])
         return row[columns[label]]
       })
 
       var plainBody = plainBodyRaw.replace(labelRegex, function (k) {
         var label = k.substring(2, k.length - 2)
         return row[columns[label]]
       })
       
       console.log('Sending Email to: ' + email)
       
       MailApp.sendEmail(email, subject, plainBody,{
         htmlBody: htmlBody,
         name: row[columns['Sender']]
       })
 
       emailMap[email] = true
 
       count++
     }
      else{
        console.log('Unable to send email to: ' + email) 
      }
    }
    else{
      console.log("Data is undefined" + data[i])
    }
  }

  ui.alert(count + ' emails sent.')
}

/*-------------------------------------
/ Forked from https://github.com/harshjv/gmail-mail-merge
/ Original Author: Harsh Vakharia
/ Modified by: David Browning
/ MIT License 
/-------------------------------------*/
