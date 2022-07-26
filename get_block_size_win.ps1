param (
[Parameter(Mandatory=$true)][string]$rg

)

###############################################
$line='='*100
Write-host $line
Write-Host "l" "DISCLAIMER".padleft(50) "l".padleft(47)
Write-Host "l `tThis script need connectivity to Azure, if you are not connected, you may get an error" -Nonewline ; Write-host "l".padleft(6)
Write-Host "l`tUse below command to login and set the subscription" -Nonewline; Write-host "l".padleft(41)
Write-Host "l `tconnect-Azaccount" -nonewline ; Write-host "l".padleft(75)
Write-Host "l `tset-Azcontext -Subscription [Subscription Name]" -nonewline ; Write-host "l".padleft(45)
Write-Host $line`n`n

$greenCheck = @{
  Object = [Char]8730
  ForegroundColor = 'Green'
  }  
################################################


Write-host "Getting VM list from resource group " -nonewline ; Write-host "$rg" -ForegroundColor Green
#Get VM list from ResourceGroup
$vmlist = Get-AzVM -Resourcegroupname $rg  -ErrorAction SilentlyContinue

if(-not $vmlist) {
	Write-Host "Error : either the RG does not exist or has no VM in it or you are not connected to Azure. Please check and try again." -BackgroundColor Red
	exit
}

Write-Host "`nFound $($vmlist.count) VM in resource group " -Nonewline
Write-Host "$rg" -ForegroundColor Green 
$count = 1
Foreach ($vm in $vmlist) {
	Write-Host "${count}. $($vm.name)"
	$count++
}

Write-Host "`nProvide the Domain login credential for connecting to VM"

#Login credentials.
$username = Read-host "Domain user name "
$password = Read-host -assecurestring "Password "

$cred = New-Object System.Management.Automation.PSCredential($username, $password)

if(-not $cred) {
	write-Host "No credentials provided to login. Script cannot continue" -backgroundColor Red
	exit
}
#Set log directory
$logdir = "$($PSScriptRoot)\output"

$tdate = get-date -Format "dd_MMM"

# check if log directory exist else create it.
if (Test-Path -Path $($logdir) ) {	$elog = "$($logdir)\BlockSizeList_$($rg)_$($tdate).xlsx" }
else { md $logdir |out-null }

#Start EXCEL
$elog = "$($logdir)\BlockSizeList_$($rg)_$($tdate).xlsx"
$excel = New-Object -ComObject excel.application
$excel.visible = $False
$excel.DisplayAlerts = $False
$workbook = $excel.Workbooks.Add()
$uregwksht= $workbook.Worksheets.Add()
$uregwksht.Name = "BlockSize"

$row = 1; $col = 1

for($i = 1; $i -le 6; $i++) {
$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
}
$uregwksht.Cells.Item($row,$col++) = "VM Name"
$uregwksht.Cells.Item($row,$col++) = "Drive Name"
$uregwksht.Cells.Item($row,$col++) = "Drive Label"
$uregwksht.Cells.Item($row,$col++) = "Drive Size (GB)"
$uregwksht.Cells.Item($row,$col++) = "BlockSize"
$uregwksht.Cells.Item($row,$col++) = "Filesystem"


	
Write-Host "`nConnecting to each VM"
foreach ($vm in $vmlist) {
	Write-Host "Getting Volume information from : " -Nonewline ; Write-host "$($vm.name) `t" -ForegroundColor Green -Nonewline ;
	$volumelist = Invoke-Command -computername $($vm.name) -Credential $cred -ScriptBlock { Get-CimInstance -ClassName Win32_Volume  | Select-Object systemname,name,Label, BlockSize,Filesystem,Capacity }
	
	if ($volumelist) {
		foreach($volume in $volumelist) {
			if($($volume.Filesystem) -eq "NTFS") {
				$row++; $col=1
				$uregwksht.Cells.Item($row,$col++) = "$($volume.systemname)"
				$uregwksht.Cells.Item($row,$col++) = "$($volume.name)"
				$uregwksht.Cells.Item($row,$col++) = "$($volume.Label)"
				$uregwksht.Cells.Item($row,$col++) =  [math]::Round($($volume.capacity)/1024/1024/1024,2)
				$uregwksht.Cells.Item($row,$col++) = "$($volume.BlockSize)"
				$uregwksht.Cells.Item($row,$col++) = "$($volume.Filesystem)"
			}
		}
	Write-Host @greenCheck
	} else { 
		Write-Host "x" -ForegroundColor Red
	}
	
}

#close and save the excel file
$workbook.Worksheets['BlockSize'].UsedRange.Columns.Autofit() | Out-Null
$workbook.WorkSheets.item("Sheet1").delete()
$workbook.Saveas($elog)  | Out-Null
$workbook.Close|Out-Null
$excel.Quit()|Out-Null

Write-Host "`nReport saved at : $elog" |Out-String -width 30
Read-Host -Prompt "`nPress any key to exit..."
