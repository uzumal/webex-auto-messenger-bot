# webex-auto-messenger-bot
This repository manages the code for a bot that periodically sends messages to a Webex chat room using Google Apps Script (GAS).

## Features

- Sends messages to a specific Webex room
- Retrieves member lists and saves them to a spreadsheet
- Counts the number of times members have received messages and distributes them evenly
- Records sent messages and selects unsent messages randomly
- Considers weekends and holidays to send messages only on weekdays
- Sets up a Google Apps Script trigger for automatic execution every morning

## Usage

- Create a Google Apps Script project and copy the script from this repository
- Modify the CONFIG variable to match your Webex room ID and spreadsheet ID
- Set the Webex access token using the setToken() function
- Run setTrigger() to set up an automated trigger at 9 AM daily

## Required Permissions

- Access to Webex API (Webex access token required)
- Edit permission for Google Spreadsheets
- Permission to retrieve Japan holiday information from Google Calendar (for holiday detection)

## Configuration (CONFIG)
```bash
const CONFIG = {
  ROOM_ID: 'YOUR_ROOM_ID',
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID',
  MESSAGE_SHEET: 'Sheet1',
  MEMBER_SHEET: 'Sheet2'
};
```
## Function List

`setToken()` - Stores the Webex access token in script properties

`send(text, personEmail = null)` - Sends a message to a specified user or everyone

`getRoomMembers()` - Retrieves room members

`updateMemberSheet(members)` - Updates the member list in the spreadsheet

`getNextMember(memberSheet)` - Selects the next member to receive a message

`clearMemberCounts(memberSheet)` - Resets message counts

`getRandomUnsentMessage(sheet)` - Retrieves a random unsent message

`sendDailyMessage()` - Selects a member and sends a message

`isWorkday(targetDate)` - Determines if a given date is a holiday

`main()` - Executes sendDailyMessage() only on weekdays

`setTrigger()` - Sets up an automated script execution trigger

## License
Distributed under the MIT License. See LICENSE for more information.