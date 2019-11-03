# VandyHacks-Auto-Auditor-with-UiPath
The goal of this project is to completely automate the process of data security checks for any organization to ensure the confidentiality and security of client data.

## What all security checks it covers
1.Antivirus status of a given machine.
2.Encryption status of data.
3.Access to storage sites such as GoogleDrive, Dropbox.
4.Access to social media platforms.
5.The number of users having access to the given machine.
6.All illegal software installed in a given machine.
7.If Bluetooth is available then the port number.
8.Track if protected data is shared by using email.

## Before you run(Assumption)
1. We have assumed that user is having given softwares installed which are Mozilla Firefox(To check stored passwords), Outlook(To send detailed report to data security team). If you dont have, You can remove this security check modules and run it.
2. Few paths are hardcoded, On the basis of organizations naming convention we can remove this assumption. Before running change the paths present in [code.vbs] file.
3. Change the email adress present at bottom of the file.

## How to run
Open main file in UiPath and run it.

## Output
Once you change the email address and run it, After few seconds you will get an automated report of the security checks performed on given local machine through email.
