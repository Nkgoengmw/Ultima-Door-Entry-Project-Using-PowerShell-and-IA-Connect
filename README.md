# Door Entry Automation Using PowerShell

## Overview

The **Door Entry Automation Project** is an end-to-end employee access provisioning solution developed using **PowerShell** and **IA-Connect**.

The solution automates the process of receiving employee provisioning requests, extracting employee information, creating employee records in the Door Entry desktop application, retrieving the generated Employee ID, and notifying the original requester.

The project was designed to reduce manual data entry, improve processing accuracy, and streamline the employee onboarding and access provisioning process.

## Key Features

- Monitors **Microsoft Outlook** for employee provisioning requests.
- Retrieves employee information from attached **Excel files**.
- Extracts and prepares employee information for processing.
- Uses **IA-Connect UI Automation (UIA)** to interact with the Door Entry desktop application.
- Automatically enters employee information into Door Entry.
- Creates new employee records.
- Captures the generated **Employee ID** from the Door Entry application.
- Sends an automated Outlook response to the original requester.
- Provides processing results and the generated Employee ID.
- Includes exception handling and logging for troubleshooting.
- Uses a modular PowerShell architecture for easier maintenance and enhancement.

## Technologies Used

| Technology | Purpose |
|---|---|
| PowerShell | Main automation and orchestration logic |
| IA-Connect | Windows desktop UI automation |
| UI Automation (UIA) | Interaction with Door Entry controls |
| Microsoft Outlook | Receiving requests and sending responses |
| Microsoft Excel | Source of employee provisioning information |
| Door Entry | Employee access provisioning application |

## Process Flow

1. **Receive Request**
   - The automation checks Outlook for employee provisioning requests.

2. **Retrieve Attachment**
   - The Excel attachment containing employee information is retrieved from the email.

3. **Extract Employee Data**
   - PowerShell reads the required employee information from the Excel file.

4. **Open Door Entry**
   - The automation connects to the Door Entry desktop application using IA-Connect.

5. **Create Employee Record**
   - Employee information is entered into the relevant fields using UI Automation commands.

6. **Capture Employee ID**
   - After the employee is successfully created, the generated Employee ID is extracted from the application.

7. **Send Response**
   - An automated email is sent to the original requester containing the processing result and generated Employee ID.

8. **Exception Handling**
   - Any processing failures are captured and logged to support troubleshooting and operational monitoring.

## Project Structure

The solution follows a modular PowerShell design. Example modules/scripts include:

```text
DoorEntry/
|-- Main.ps1
|-- ReadEmail.ps1
|-- DoorEntry.ps1
|-- Send-EmployeeResponse.ps1
|-- README.md
```

### `Main.ps1`

Controls the overall process and coordinates the different PowerShell modules.

### `ReadEmail.ps1`

Responsible for retrieving Outlook requests and extracting information from the attached Excel file.

### `DoorEntry.ps1`

Uses IA-Connect UI Automation commands to interact with the Door Entry desktop application, create the employee record, and retrieve the generated Employee ID.

### `Send-EmployeeResponse.ps1`

Sends the processing result and Employee ID back to the original requester.

## Error Handling

The solution includes error handling using PowerShell exception handling mechanisms such as `try`, `catch`, and `finally`.

Typical exceptions handled include:

- Missing or invalid Excel attachments
- Missing employee information
- Door Entry application unavailable
- UI elements not found
- Employee creation failure
- Employee ID extraction failure
- Outlook email processing failure

## Benefits

- Reduces manual employee provisioning activities.
- Minimizes data-entry errors.
- Improves turnaround time for Door Entry requests.
- Provides automated requester notifications.
- Creates a repeatable and standardized provisioning process.
- Modular PowerShell structure makes the solution easier to maintain and extend.

## Future Enhancements

Potential enhancements include:

- Centralized logging and reporting.
- Improved employee-data validation before processing.
- Automated retry mechanisms for temporary UI failures.
- Processing dashboards and operational metrics.
- Configuration files for environment-specific settings.
- More detailed audit logging for processed requests.

## Author

**Matome N**  
Developer
