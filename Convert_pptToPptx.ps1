$PpFixedFormat = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLPresentation
write-host $PpFixedFormat
$PowerPoint = New-Object -ComObject PowerPoint.application
$folderpath = "C:\Docs\Ppt to pptx.ppt"
$filetype ="*ppt"
Get-ChildItem -Path $folderpath -Include $filetype -recurse |
ForEach-Object `
{
    $path = ($_.fullname).substring(0, ($_.FullName).lastindexOf("."))
 
    "Converting $path"
    $Presentation = $PowerPoint.Presentations.open($_.fullname)
 
    $path += ".pptx"
    $Presentation.saveas($path, $PpFixedFormat)
    $Presentation.close()
 
    $oldFolder = $path.substring(0, $path.lastIndexOf("\")) + "\old"
 
    write-host $oldFolder
    if(-not (test-path $oldFolder))
    {
        new-item $oldFolder -type directory
    }
 
    move-item $_.fullname $oldFolder
 
}
$PowerPoint.Quit()
$PowerPoint = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()