# Automated-Finance-Claims-Management-System
Automated Finance Claims Management System for Raffles Hall Finance


## INSTALLATION (One-Time Setup):
1. Create a brand new Google Sheet
2. Extensions > Apps Script
3. Delete all default code
4. Paste the entire automatedClaims.gs
5. Save (Ctrl/Cmd + S)
6. Click Run > select "setupClaimsSystem"
7. Grant permissions when prompted
8. Fill in the Config sheet that appears
9. Enable Drive API: Click + next to Services > Add "Drive API" v3
10. Move the Google Sheets into the Automated Claims folder
11. Copy the Template RFP and Summary into the Automated Claims folder
12. Set them to 'Anyone with the link can view'
13. Copy their links and extract the ID and input into the Config sheet

Example 1: https://docs.google.com/spreadsheets/d/1DyJGAGqTKjra-2ErZ885EEU2N-wVJetvL82kNQYS9Y/edit?usp=drive_link

ID: 1DyJGAGqTKjra-2ErZ885EEU2N-wVJetvL82kNQYS9Y

Example 2: https://docs.google.com/document/d/1qvpOijRMO5chlJIsCYLpvVVmLrawgHPn0c7uRhPBUo/edit?usp=drive_link

ID: 1qvpOijRMO5chlJIsCYLpvVVmLrawgHPn0c7uRhPBUo

14. Hide the Config sheet
15. Copy the Finance Form into the Automated Claims folder
16. Link the Finance Form to the Google Sheets (Link to existing spreadsheet and select the Google Sheets created)
17. Scroll to the bottom of the Form Responses 1 Sheet and add 3000 more rows
18. Add 2 columns to the left of the Form Responses 1 Sheet
19. Select the 2 columns except the headers and click Insert -> Checkboxes
20. Rename the headers Processed and Error
21. You are done :D
