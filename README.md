# MLSA-Event-Certificate-Automation
Being a Microsoft Learn Student Ambassador, it's our duty to share the knowledge by hosting events. And the people who attend the event earn a Microsoft Learn Student Ambassador Certificate of Participation. Manually making a hundred certificates and mailing them all can be tedious. 

But this script has got your back and can be used to generate and mail hundreds of certificates to the respective person within few seconds!

## Steps to be followed

### 1. Create a template and save in the working directory
Download the certificate template from Microsoft Teams. Type "Here" where you want the name of the certificate holder. Insert the event's name and your name in the respective places
<img src="https://github.com/chiragjagad/MLSA-Event-Certificate-Automation/blob/master/readme/template.png">
### 2. Make list.csv
Add names and mails in the list.csv file
### 3. Make changes in script.py
-> Add your path of working directory

-> Add your email address password

-> Add body and subject of the mail
### 4. Create two folders
Create two folders, 'certificates-pdf' and 'certificates-word', to store all the certificates in pdf and word format, respectively
### 4. Run the script
-> After running the script, all the certificates would be saved in the respective folder according to their format

-> The certificate would be saved as 'Name.pdf' and 'Name.docx'

-> The PDF file of the certificate would be sent as an attachment in the mail to the respective person. The mail would also contain the subject and the body of the mail.
### 5. Final Result!
John Doe.pdf
<img src="https://github.com/chiragjagad/MLSA-Event-Certificate-Automation/blob/master/readme/certificate.png">

