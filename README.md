# XML to SAP Using Power Automate Desktop

This project allows you to upload some files to SAP ERP by using Power Automate Desktop, using a main flow that interacts with an Outlook instance and performs actions based on the selection of different "Usinas" (steel plants). This automation can be useful for managing tasks or processes specific to each steel plant, improving efficiency, and ensuring consistency in operations.

Below is an explanation of each section of the script.

## The main flow
Allows the user to choose between several options of Steel Plants presenting a selection dialog to the user, and calls specific sub-flows accordingly. For each Steel Plants, we have subflows that works on it, by using several lines of automation and specific keywords.

## Sub-Flows

The script automates a process using Microsoft Power Automate Desktop, handling email attachments, XML file processing, and interactions with SAP and Excel. For each Sub-Flows it firsts obtains emails and saves attachments to a specific folder. The script prompts the user to enter a date range to process the files and finds the corresponding rows in Excel. If no match is found, it asks the user to re-enter the dates until matches are found. The script then processes the XML files within the selected date range, moving them to a temporary folder. It logs into SAP, starts specific transactions, fills in necessary fields using keyboard commands, and executes the transactions. Finally, it empties the original folder where the XML files were saved and closes the Excel and SAP instances.

This workflow saves time and reduces errors by automating the processing of XML files, Excel interactions, and SAP operations.
