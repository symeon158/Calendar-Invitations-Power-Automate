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

![Flow Diagram]("C:\Users\sy.papadopoulos\OneDrive - Alumil S.A\Pictures\Screenshots\Calendar Invites.png")

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

