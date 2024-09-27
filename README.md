# Automate Induction Training

## Overview

This project automates the process of creating and sending calendar invitations for induction training sessions. Using dynamic fields extracted from Excel, the flow leverages Office Scripts to trigger the workflow and Microsoft Power Automate (Flow) to create calendar events efficiently.

---

## Table of Contents

1. [Overview](#overview)
2. [Features](#features)
3. [Flow Diagram](#flow-diagram)
4. [Power Automate Flow Details](#power-automate-flow-details)
5. [Office Script for Excel](#office-script-for-excel)
6. [Setup Instructions](#setup-instructions)
7. [Usage](#usage)
8. [Screenshots](#screenshots)
9. [License](#license)

---

## Features

- **Automated Event Creation**: Automatically creates calendar events for induction training based on data in an Excel sheet.
- **Dynamic Invitations**: Sends calendar invites with dynamic details like **Training Topic**, **Start and End Times**, **Attendee Emails**, and **Location**.
- **Customizable Content**: The invitation body can be personalized with placeholders for **Colleague Name**, **Job Title**, and **Meeting Place**.
- **Manual Trigger Option**: Easily initiate the process through an Office Script button inside Excel.
- **Predefined Categories & Location**: Add categories and locations to events for easy calendar management.

---

## Flow Diagram

### Power Automate Process

Below is a flow diagram of the Microsoft Power Automate process that handles the creation of calendar invitations.

1. **Manual Trigger**: The process is manually initiated through an Excel button.
2. **List Rows Present in a Table**: Retrieves all rows from an Excel table containing the necessary data.
3. **Apply to Each Row**: Iterates over each row to apply conditional logic.
4. **Condition Check**: Evaluates whether an event should be created based on a condition.
5. **Create Event (V4)**: Creates a calendar event with the details provided.
---
![Calendar Invites](https://github.com/user-attachments/assets/73e1116d-aebf-4f43-9bee-66fb639460fc)
![create event1](https://github.com/user-attachments/assets/17812ad6-bd62-47ea-bcc6-da11ca19145c)
![create events2](https://github.com/user-attachments/assets/2df50822-0a56-4fe4-b74e-a51713fd48af)


---

## Power Automate Flow Details

The flow uses the following steps to automate the calendar invitation process:

1. **Manual Trigger**: This step initiates the flow when manually triggered from Excel.
2. **List Rows Present in a Table**: Fetches all rows from the table in the Excel workbook. It accesses data such as **DateTimeS**, **DateTimeE**, **Training Topic**, **Email**, etc.
3. **Apply to Each Row**: Iterates through each row in the table to perform actions based on the data.
4. **Condition Check**: Applies conditions to determine whether an event should be created. If the condition is `True`, it proceeds to create a calendar event.
5. **Create Event (V4)**: Uses details from each row in Excel to create a calendar event in Microsoft Outlook. The parameters include:
   - **Calendar Id**: Specifies the calendar where the event is created.
   - **Subject**: Combines "Induction Training" with the **Training Topic** from the Excel data.
   - **Start & End Times**: Uses **DateTimeS** and **DateTimeE** from the Excel table.
   - **Time Zone**: Configured to the appropriate time zone (e.g., `(UTC+02:00) Athens, Bucharest`).
   - **Required Attendees**: Populates with emails from the **Email** column.

---

## Office Script for Excel

The following Office Script enables the manual triggering of the flow from within the Excel workbook:

```typescript
function main(workbook: ExcelScript.Workbook) {
    let httpRequest = new XMLHttpRequest();
    let myPath = workbook.getActiveWorksheet().getRange("AB1").getText();
    httpRequest.open("GET", myPath, false);
    httpRequest.send(null);
}

---

Setup Instructions
Prerequisites
Microsoft Power Automate (Flow): To create the automation flow.
Microsoft Excel with Office Scripts enabled: For running the script that triggers the flow.
Microsoft Outlook/Exchange Calendar: To send and manage calendar invitations.
Steps
Create the Power Automate Flow:

Manual Trigger: Start the flow with a manual trigger for easy initiation.
List Rows Present in a Table: Add an action to list rows from the Excel table containing event data.
Apply to Each and Condition Check: Loop through each row to evaluate conditions for creating events.
Create Event (V4): Use this action to create events in the Outlook/Exchange Calendar when conditions are met.
Set up the Excel Workbook:

Ensure your Excel table has columns like DateTimeS (Start Time), DateTimeE (End Time), Email (Attendees), Meeting Place, etc.
Add a button to the workbook and link it to the provided Office Script to trigger the flow manually.
Configure the Office Script:

Insert the Office Script in Excel to automate the flow trigger.
Use the script button to initiate the process whenever necessary.
Usage
Open the Excel workbook containing the induction training data.
Click the button linked to the Office Script to trigger the Power Automate flow.
The flow will loop through each row of the Excel table, creating a calendar event for each training session that meets the specified condition.
Screenshots
Power Automate Flow Structure
This flow represents the logic used to automate the calendar invitation creation process. The process is straightforward and ensures that each row in the Excel table is evaluated correctly.


Create Event Action Details
The Create Event (V4) action contains various parameters such as Calendar Id, Subject, Start Time, End Time, Time Zone, and Required Attendees.


Calendar Id: Specifies the calendar where the event is created.
Subject: Sets the event topic, combining "Induction Training" with the training topic from the Excel data.
Start Time & End Time: Uses dynamic fields (DateTimeS and DateTimeE) from the Excel table to schedule the event.
Time Zone: Configured to the appropriate region (e.g., (UTC+02:00) Athens, Bucharest).
Required Attendees: Automatically populates attendees' emails from the Excel data.
Calendar Event Body Template
The body of the invitation is formatted to include the relevant details for the induction session.

