# -Email-Automation-with-VBA
1. Email Automation with VBA:
We developed a VBA macro to send automated reminder emails for critical tasks.
The macro filters rows based on unique Event IDs in Column N (EV Event ID), processes only visible rows, and creates emails for each unique Event ID.
2. Email Content:
The email content includes a table with information about critical tasks, including:
Days Remaining (Column B)
Site Name (Column E)
Event Type (Column G)
Contract Type (Column I)
GTM Planning ID (Column M)
Event ID (Column N)
Activity ID (Column O)
Activity Title (Column Q)
Task Due Date (Column T)
CPM Name (Column AL)
CPM SSO (Column AM)
3. Email Recipients:
To Field: Email addresses from Column AM (CPM SSO) are placed in the To field.
CC Field: Email addresses from Column AS are placed in the CC field.
4. Handling Filters and Data Display:
The code filters the data by Event ID (Column N), ensuring that only relevant rows are processed.
The email is created dynamically, and the table displays task details for each filtered event.
5. Outlook Integration:
The script uses Outlook to create and display emails, automatically pulling data into the HTMLBody of the email.
The signature from Outlook is retained by displaying the email first, capturing the default signature, and appending it to the custom email content.
6. Error Handling and User Interface:
If no visible rows are found for an Event ID, a message is shown to alert the user.
The macro closes the workbook after processing and provides a success message after the emails are created.
7. Workflow Execution:
The user is prompted to select an Excel file, and the macro processes the data accordingly.
It ensures that email notifications for tasks are generated with the correct recipient details based on the Event ID.
This process automates the sending of reminder emails for critical tasks based on data in the Excel file, improving efficiency and consistency.
