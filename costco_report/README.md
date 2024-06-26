# This program is to be used to generate reports.
### How does it work?
- Email needs OAUTH approval, otherwise the program will not work.
- Once approved, the "SEND_TO" constant must be changed to the email you'd like the report to be sent to.
- The program will grab the most recent email of a specific subject criteria, download and unzip the attachment in-memory and perform data transformation into a formatted report and then saved into a csv object that is emailed to the corresponding email fron SEND_TO
- If your token expires, running the program should prompt you to re-select your Email for OAUTH, which will then regenerate a new token
