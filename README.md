# ScrubEmailAddressinOutlook
Scrub your inbox, sent folder, and calendars for unique email addresses. Can choose to ignore internal domains 

Date Range Input:
 - The code prompts the user to input the start and end dates via a single input box.
 - The input should be in the format MM/DD/YYYY - MM/DD/YYYY.
 - The input is split into start and end dates, and these are used to filter emails and calendar items within the specified date range.

Domain Exclusion:
 - The script asks the user whether they want to exclude *@XXXXX.com or specify other domains to exclude.
 - The input is split into an array of domains, and the script checks if any email address matches these domains before adding it to the Excel file.
 - If no exclusion domains are provided, the script does not exclude any emails, i.e., includes all email addresses.
   
How to Run:
 - Open Outlook
 - Ensure Macros are enabled
 - Press Alt + F11 to open the VBA editor
 - Insert a new module and paste the code into it
 - Run the ScrubEmailAddresses macro
 - Follow the prompts to input the date range and exclusion domains
