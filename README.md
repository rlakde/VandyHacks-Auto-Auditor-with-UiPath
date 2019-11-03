# VandyHacks-Auto-Auditor-with-UiPath
The goal of this project is to completely automate the process of data security checks for any organization to ensure the confidentiality and security of client data.

## What all security checks it covers
<b>1.</b> Antivirus status of a given machine.<br/>
<b>2.</b> Encryption status of data.<br/>
<b>3.</b> Access to storage sites such as GoogleDrive, Dropbox.<br/>
<b>4.</b> Access to social media platforms.<br/>
<b>5.</b> The number of users having access to the given machine.<br/>
<b>6.</b> All illegal software installed in a given machine.<br/>
<b>7.</b> If Bluetooth is available then the port number.<br/>
<b>8.</b> Track if protected data is shared by using email.<br/>

## Before you run(Assumptions)
<b>1.</b> We have assumed that user is having given softwares installed which are Mozilla Firefox(To check stored passwords), Outlook(To send detailed report to data security team). If you dont have, You can remove this security check modules and run it.<br/>
<b>2.</b> Few paths are hardcoded, On the basis of organizations naming convention we can remove this assumption. Before running change the paths present in [code.vbs](https://github.com/rlakde/VandyHacks-Auto-Auditor-with-UiPath/blob/master/code.vbs) file.<br/>
<b>3.</b> Change the email adress present at bottom of the file.<br/>

## How to run
Open main file in UiPath and run it.

## Output
Once you change the email address and run it, After few seconds you will get an automated report of the security checks performed on given local machine through email.
