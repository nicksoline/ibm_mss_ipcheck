#Input a source excel file with list of IP's in its first sheet in column A
$ipsource = Read-Host -Prompt "`r`nInput path to IP source file`r`n"
Write-Host "`r`n"$ipsource
#Open source file
$Excel = New-Object -ComObject Excel.Application
$Workbook = $Excel.Workbooks.Open($ipsource)
$Worksheet = $Workbook.Sheets.Item(1)
Write-Host "`r`n"$Worksheet.Name"`r`n"
#get data from source
$rowmax = ($Worksheet.UsedRange.Rows).count
$rowA,$colA = 1,1
$R = Invoke-WebRequest https://www.ipvoid.com/ip-blacklist-check/
$Form = $R.Forms[0]
$Form.fields
$Form.fields["ipAddr"] = 255.255.255.255
$Form.fields
function global:ipcheck {
    param ($ip)
    $Form.fields["ipAddr"] = $ip
    $Form.fields
    $Uri = "https://www.ipvoid.com/ip-blacklist-check/" + $Form.Action
    $R = Invoke-WebRequest -Uri $Uri -Method POST -Body $Form.fields
    $R.StatusDescription
    return
}
for ($i=0; $i -le $rowmax-1; $i++)
{
    $ip = $Worksheet.Cells.Item($rowA+$i, $colA).text
    global:ipcheck($ip)
}
$Excel.quit()