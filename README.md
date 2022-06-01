Anything email related  
  
new-mtastsreport.ps1 - Create CSV report from JSON files sent to rua email address in _smtp._tls record.  
* Prompts for Outlook folder and saves all attachments from selected folder.  
* Extract JSON files from GZ attachments using 7Zip.  
* Create a CSV report from the JSON files.

Fix-O365ProxyAddress.ps1 - Fix proxy address on hybrid mailboxes that don't have email policy inheritance enabled.  
* Enable inheritance and wait 1 second  
* Disable inheritance and wait 1 second  
* Set the primary email address back to what it was originally