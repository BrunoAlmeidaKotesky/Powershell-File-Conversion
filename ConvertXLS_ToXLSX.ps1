$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook
write-host $xlFixedFormat
$excel = New-Object -ComObject excel.application
$excel.visible = $true
$folderpath = "C:\Docs\excelEx.xls"
$filetype ="*xls"
Get-ChildItem -Path $folderpath -Include $filetype -recurse | 
ForEach-Object `
{
	$path = ($_.fullname).substring(0, ($_.FullName).lastindexOf("."))
	s
	"Converting $path"
	$workbook = $excel.workbooks.open($_.fullname)

	$path += ".xlsx"
	$workbook.saveas($path, $xlFixedFormat)
	$workbook.close()
	
	$oldFolder = $path.substring(0, $path.lastIndexOf("\")) + "\old"
	
	write-host $oldFolder
	if(-not (test-path $oldFolder))
	{
		new-item $oldFolder -type directory
	}
	
	move-item $_.fullname $oldFolder
	
}
$excel.Quit()
$excel = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()