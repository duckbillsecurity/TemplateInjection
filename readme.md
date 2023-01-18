## RemoteTemplateInjection

Injects Microsoft template with macro into a Microsoft Word document.  

### Requirements

PowerShell version 5 or greater.

### Usage

### Basic steps

1. Create a Microsoft Word template with macro, and save as a .dotm file (Word Macro-Enabled Template (*.dotm)
2. Rename above to extension .dot (Word 97-2003 Template).
3. Create a new document from this template, and save as a Word .docx file.
4. Move the .dot file to a remote web server.
5. Use the PowerShell script to inject reference of the remote template location into the .docx file.

### References

- https://www.ired.team/offensive-security/initial-access/phishing-with-ms-office/inject-macros-from-a-remote-dotm-template-docx-with-macros
- https://github.com/JohnWoodman/remoteinjector
- https://support.microsoft.com/en-us/topic/a-potentially-dangerous-macro-has-been-blocked-0952faa0-37e7-4316-b61d-5b5ed6024216




