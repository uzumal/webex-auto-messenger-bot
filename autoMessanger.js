const CONFIG = {
  ROOM_ID: 'XXX', // Room ID you want to send messages to
  SPREADSHEET_ID: 'XXX', // Spreadsheet ID
  MESSAGE_SHEET: 'Sheet1', // Sheet name for messages
  MEMBER_SHEET: 'Sheet2' // Sheet name for members
};

function setToken() {
  PropertiesService.getScriptProperties().setProperty(
    'WEBEX_TOKEN', 
    'XXX' // Webex token
  );
}

function send(text, personEmail = null) {
  const messageText = personEmail ? `<@personEmail:${personEmail}>さん\n${text}` : text;
  
  const data = {
    'roomId': CONFIG.ROOM_ID,
    'markdown': messageText
  };
  
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': { 
      'Authorization': 'Bearer ' + PropertiesService.getScriptProperties().getProperty('WEBEX_TOKEN')
    },
    'payload': JSON.stringify(data),
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch('https://api.ciscospark.com/v1/messages', options);
    const responseCode = response.getResponseCode();
    Logger.log('Message send response code: ' + responseCode);
    Logger.log('Message content: ' + messageText);
    
    if (responseCode !== 200) {
      Logger.log('Error response: ' + response.getContentText());
      return false;
    }
    return true;
  } catch (error) {
    Logger.log('Error sending message: ' + error);
    return false;
  }
}

function getRoomMembers() {
  const options = {
    'method': 'get',
    'headers': { 
      'Authorization': 'Bearer ' + PropertiesService.getScriptProperties().getProperty('WEBEX_TOKEN')
    },
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(`https://api.ciscospark.com/v1/memberships?roomId=${CONFIG.ROOM_ID}`, options);
    const responseCode = response.getResponseCode();
    Logger.log('Get members response code: ' + responseCode);
    
    if (responseCode === 200) {
      const members = JSON.parse(response.getContentText()).items;
      const humanMembers = members.filter(member => !member.personEmail.endsWith('@webex.bot'));
      Logger.log('Human members: ' + JSON.stringify(humanMembers));
      return humanMembers;
    } else {
      Logger.log('Error response: ' + response.getContentText());
      return [];
    }
  } catch (error) {
    Logger.log('Error getting members: ' + error);
    return [];
  }
}

function updateMemberSheet(members) {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const memberSheet = spreadsheet.getSheetByName(CONFIG.MEMBER_SHEET);
  
  if (memberSheet.getLastRow() < 1) {
    memberSheet.getRange(1, 1, 1, 3).setValues([['Email', 'Name', 'Count']]);
  }
  
  const currentData = memberSheet.getDataRange().getValues();
  const existingEmails = currentData.slice(1).map(row => row[0]);
  
  members.forEach(member => {
    const index = existingEmails.indexOf(member.personEmail);
    if (index === -1) {
      memberSheet.appendRow([member.personEmail, member.personDisplayName, 0]);
    }
  });
  
  return memberSheet;
}

function getNextMember(memberSheet) {
  const data = memberSheet.getDataRange().getValues();
  const members = data.slice(1);
  
  if (members.length === 0) return null;
  
  const minCount = Math.min(...members.map(m => m[2]));
  const candidateMembers = members.filter(m => m[2] === minCount);
  
  const selected = candidateMembers[Math.floor(Math.random() * candidateMembers.length)];
  return {
    email: selected[0],
    name: selected[1],
    count: selected[2],
    row: data.findIndex(row => row[0] === selected[0]) + 1
  };
}

function clearMemberCounts(memberSheet) {
  const lastRow = memberSheet.getLastRow();
  if (lastRow > 1) {
    memberSheet.getRange(2, 3, lastRow - 1, 1).setValue(0);
  }
}

function getRandomUnsentMessage(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  
  const unsentRows = [];
  for(let i = 2; i <= lastRow; i++) {
    if(!sheet.getRange(i, 4).getValue()) {
      const message = sheet.getRange(i, 1).getValue();
      if (message) unsentRows.push(i);
    }
  }
  
  if (unsentRows.length === 0) {
    if (lastRow > 1) sheet.getRange(2, 4, lastRow - 1).clearContent();
    return null;
  }
  
  const randomRowIndex = Math.floor(Math.random() * unsentRows.length);
  return {
    row: unsentRows[randomRowIndex],
    message: sheet.getRange(unsentRows[randomRowIndex], 1).getValue()
  };
}

function sendDailyMessage() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const messageSheet = spreadsheet.getSheetByName(CONFIG.MESSAGE_SHEET);
    if (!messageSheet) return;
    
    const members = getRoomMembers();
    if (members.length === 0) return;
    
    const memberSheet = updateMemberSheet(members);
    const selectedMember = getNextMember(memberSheet);
    if (!selectedMember) return;
    
    const messageInfo = getRandomUnsentMessage(messageSheet);
    if (!messageInfo) {
      Logger.log('All messages have been sent. Clearing member counts.');
      clearMemberCounts(memberSheet);
      return;
    }
    
    if(send(messageInfo.message, selectedMember.email)) {
      messageSheet.getRange(messageInfo.row, 2).setValue(new Date());
      messageSheet.getRange(messageInfo.row, 3).setValue(selectedMember.email);
      messageSheet.getRange(messageInfo.row, 4).setValue(true);
      
      memberSheet.getRange(selectedMember.row, 3).setValue(selectedMember.count + 1);
      Logger.log(`Message sent to: ${selectedMember.name} (${selectedMember.email})`);
    }
  } catch (error) {
    Logger.log('Error in sendDailyMessage: ' + error);
  }
}

function isWorkday(targetDate) {
  const dayOfWeek = targetDate.getDay();
  if(dayOfWeek === 0 || dayOfWeek === 6) return false;
  
  try {
    const calJpHoliday = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
    if(calJpHoliday.getEventsForDay(targetDate).length > 0) return false;
  } catch (error) {
    Logger.log('Error checking holiday: ' + error);
    return true;
  }
  
  return true;
}

function main() {
  try {
    const today = new Date();
    if(isWorkday(today)) sendDailyMessage();
  } catch (error) {
    Logger.log('Error in main: ' + error);
  }
}

function setTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    
    ScriptApp.newTrigger('main')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
  } catch (error) {
    Logger.log('Error setting trigger: ' + error);
  }
}