Add-Type -assembly "system.io.compression.filesystem"

$source = 'C:\dev\excel-plugin-demo\demo-plugin\demo-plugin\bin\Debug'    
$destinationPath = "C:\dev\excel-plugin-demo\deploy"
$destination = $destinationPath + "\excel-plugin.zip"

If(Test-path $destinationPath) 
{
    "Cleaning up"
    if (Test-Path $destination) {Remove-item $destination}
}
Else
{
    "Creating destination folder"
    md $destinationPath
}

[io.compression.zipfile]::CreateFromDirectory($Source, $destination)
"Zip complete"