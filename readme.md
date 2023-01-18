## RemoteTemplateInjection

Injects Microsoft template with macro into a Microsoft Word document.  

### Requirements

PowerShell version 5 or greater.

### Usage

### Basic steps

1. Create a Microsoft Word template with macro, and save as a .dot file (Word 97-2003 Template).
2. Create a new document from this template, and save as a Windows Word .docx file.
3. Move the .dot file to a remote web server.
4. Use the PowerShell script to inject reference of the remote template location into the .docx file.

### References

- https://www.ired.team/offensive-security/initial-access/phishing-with-ms-office/inject-macros-from-a-remote-dotm-template-docx-with-macros
- https://github.com/JohnWoodman/remoteinjector




