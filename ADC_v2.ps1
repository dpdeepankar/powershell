#################################################################################
#################################################################################
##                                                                             ##
##  Azure Data Collector v2                                                    ##
##  v1.0 D.P initial version                                                   ##
##  v1.1 D.P added UDR, NSG, ASG and Firewall                                  ##
##  v2.0 D.P complete re-write, pending to add UDR and Firewall definitions.   ##
##                                                                             ##
#################################################################################
#################################################################################

# Script current version
$version = "v2.0"

Function logger {
	param ($msg, $t)
	
	if ($t -eq 'i') { 
	Write-Host "`t[INFO] " -ForegroundColor Green -NoNewline
	Write-Host "$msg" -ForegroundColor White
	}
	else { 
		Write-Host "[Error] " -ForegroundColor Red -NoNewline
		Write-Host "$msg" -ForegroundColor Red
	}
	
}



Function ascii {
	clear-host
	Write-host ""
	Write-Host "`t:::::::::'###::::'########:'##::::'##:'########::'########::::"
	Write-Host "`t::::::::'## ##:::..... ##:: ##:::: ##: ##.... ##: ##.....:::::"
	Write-Host "`t:::::::'##:. ##:::::: ##::: ##:::: ##: ##:::: ##: ##::::::::::"
	Write-Host "`t::::::'##:::. ##:::: ##:::: ##:::: ##: ########:: ######::::::"
	Write-Host "`t:::::: #########::: ##::::: ##:::: ##: ##.. ##::: ##...:::::::"
	Write-Host "`t:::::: ##.... ##:: ##:::::: ##:::: ##: ##::. ##:: ##::::::::::"
	Write-Host "`t:::::: ##:::: ##: ########:. #######:: ##:::. ##: ########::::"
	Write-Host "`t::::::..:::::..::........:::.......:::..:::::..::........:::::"
	
	Write-Host "`n`t================== "-NoNewline; Write-Host "AZURE DATA COLLECTOR $version" -ForegroundColor Green -NoNewline; Write-Host " ==================`n"
	
}

Function get_SubscriptionList {
	Write-host "`n`t`tList of Available Subscriptions: `n"
	$global:subs = Get-Azcontext -listAvailable
	
	$count=0
	foreach ($name in $global:subs.Subscription.Name)
		{
			$count++
			Write-host "`t`t${count}. $name"
		}

}

Function get_ResourceGroup {
	if($global:mode -ne 15) { ascii }
	logger -t i -msg "Collecting Resource Group details..."
	
	$RGS = Get-AzResourceGroup
	$datafound=0
	#excel creation
	$row = 1; $col = 1
	$uregwksht= $workbook.Worksheets.item(1)
	$uregwksht.Name = 'ResourceGroup'
	
	if ($RGS)
	{
		for($i = 1; $i -le 3; $i++) 
		{
			$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
			$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}

		$uregwksht.Cells.Item($row,$col++) = "Resource Group Name"
		$uregwksht.Cells.Item($row,$col++) = "Location"
		$uregwksht.Cells.Item($row,$col++) = "Subscription"
		
		foreach($rg in $RGS)
		{
			$row++; $col = 1
			$uregwksht.Cells.Item($row, $col++) = "$($rg.ResourceGroupName)".tolower().tolower()
			$uregwksht.Cells.Item($row, $col++) = "$($rg.Location)".tolower().tolower()
			$uregwksht.Cells.Item($row, $col++) = "$($global:subs.Subscription.Name[$($global:sel_subs -1)])".tolower().tolower()
			$datafound=1
		}
	
	if(-not $datafound){ $row=2;$col=1 ;$uregwksht.Cells.Item($row,$col) = "This subscription do not have any Resource groups"	}

	#Auto-fit all coumns as per the content length
	$workbook.Worksheets['ResourceGroup'].UsedRange.Columns.Autofit() | Out-Null

	}
	else
	{
		$uregwksht.Cells.Item($row+1,$col) = "This subscription do not have a Resource Group"
	}
}

Function get_AvailabilitySet {
	if($global:mode -ne 15) { ascii }
	logger -t i -msg "Collecting Availability Sets details..."
	$avsets = Get-AzAvailabilitySet
	$datafound=0
	
	if ($avsets){
		#excel creation
		$row = 1; $col = 1
		$uregwksht= $workbook.Worksheets.Add()
		$uregwksht.Name = 'Availability_Set'

		for($i = 1; $i -le 8; $i++) {
			$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
			$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}
		
		$uregwksht.Cells.Item($row,$col++) = "Availability Set Name"
		$uregwksht.Cells.Item($row,$col++) = "VM"
		$uregwksht.Cells.Item($row,$col++) = "Fault Domain"
		$uregwksht.Cells.Item($row,$col++) = "Update Domain"
		$uregwksht.Cells.Item($row,$col++) = "Resource Group"
		$uregwksht.Cells.Item($row,$col++) = "PPG Name"
		$uregwksht.Cells.Item($row,$col++) = "Location"
		$uregwksht.Cells.Item($row,$col++) = "Subscription"
		
		foreach($avset in $avsets){
			$start_row = $row+1
			
			if ($($avset.VirtualMachinesReferences)) {
				foreach($VM in $($avset.VirtualMachinesReferences)) {
					$row++; $col = 1
					$uregwksht.Cells.Item($row, $col++) = "$($avset.Name)".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$($VM.id.split('/')[8])".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$($avset.PlatformFaultDomainCount)".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$($avset.PlatformUpdateDomainCount)".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$($avset.ResourceGroupName)".tolower()
					if($($avset.proximityPlacementGroup) ){ $uregwksht.Cells.Item($row, $col++) = "$($avset.proximityPlacementGroup.id.split('/')[8])".tolower() }
					else { $uregwksht.Cells.Item($row, $col++)  = "NA"}
					$uregwksht.Cells.Item($row, $col++) = "$($avset.location)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($global:subs.Subscription.Name[$($global:sel_subs -1)])".tolower()
					$datafound=1
				}
			} 
			else {
				$row++; $col = 1
				$uregwksht.Cells.Item($row, $col++) = "$($avset.Name)".tolower()
				$uregwksht.Cells.Item($row, $col++) = "NA"
				$uregwksht.Cells.Item($row, $col++) = "$($avset.platformFaultDomainCount)".tolower()
				$uregwksht.Cells.Item($row, $col++) = "$($avset.platformUpdateDomainCount)".tolower()
				$uregwksht.Cells.Item($row, $col++) = "$($avset.resourceGroupname)".tolower()
				$uregwksht.Cells.Item($row, $col++) = "$($avset.proximityPlacementGroup.id.split('/')[8])".tolower()
				$uregwksht.Cells.Item($row, $col++) = "$($avset.location)".tolower()
				$uregwksht.Cells.Item($row,$col++) = "$($global:subs.Subscription.Name[$($global:sel_subs -1)])".tolower()
				$datafound=1
			}
			
			$end_row = $row
			
			function merge_cell ($col, $srow, $erow) 
			{
			$MergeCells = $uregwksht.Range("$col${srow}:$col$erow")
			$MergeCells.Select() |Out-Null
			$MergeCells.MergeCells = $true  
			}
			
			if ($($avset.VirtualMachinesReferences.count) -gt 1) {
					merge_cell "A" $start_row $end_row
					merge_cell "C" $start_row $end_row
					merge_cell "D" $start_row $end_row
					merge_cell "E" $start_row $end_row
					merge_cell "F" $start_row $end_row
					merge_cell "G" $start_row $end_row
					merge_cell "H" $start_row $end_row
				}
		
		}

		$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
		$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)
		$workbook.Worksheets['Availability_Set'].UsedRange.Columns.Autofit() | Out-Null
	
	}

		if(-not $datafound){ $row=2;$col=1 ;$uregwksht.Cells.Item($row,$col) = "This subscription do not have any Availability sets"	}


}

Function get_PPG {
	if($global:mode -ne 15) { ascii }
	logger -t i -msg "Collecting Proximity Placement Group details..."
	$PPGs = Get-AzProximityPlacementGroup
	$datafound=0
	#excel creation
	$row = 1; $col = 1
	$uregwksht= $workbook.Worksheets.Add()
	$uregwksht.Name = 'Proximity Placement Group'
	
	if($PPGs) {
		for($i = 1; $i -le 6; $i++) {
			$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
			$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}
		
		$uregwksht.Cells.Item($row,$col++) = "PPG Name"
		$uregwksht.Cells.Item($row,$col++) = "Location"
		$uregwksht.Cells.Item($row,$col++) = "PPG Type"
		$uregwksht.Cells.Item($row,$col++) = "Resource Group"
		$uregwksht.Cells.Item($row,$col++) = "Availability Set"
		$uregwksht.Cells.Item($row,$col++) = "Member VMs"
		
		foreach($ppg in $PPGs)
		{
			$row++; $col = 1		
			$uregwksht.Cells.Item($row, $col++) = "$($ppg.name)".tolower()
			$uregwksht.Cells.Item($row, $col++) = "$($ppg.location)".tolower()
			$uregwksht.Cells.Item($row, $col++) = "$($ppg.proximityPlacementGroupType)".tolower()
			$uregwksht.Cells.Item($row, $col++) = "$($ppg.resourceGroupname)".tolower()
			$member_avset=""
			foreach($avset in $($ppg.AvailabilitySets) ){
				if($member_avset -ne "") {
					$member_avset = $member_avset + ",`n" + $($avset.id.split('/')[8])
				} else { $member_avset = $($avset.id.split('/')[8]) }
			}
			$uregwksht.Cells.Item($row, $col++) = "$member_avset".tolower()
			
			$member_vm = "" ; $count=0
			foreach($vm in $($ppg.VirtualMachines) ) {
				if($member_vm -ne "" ) {
					$member_vm = $member_vm + "," + $($vm.id.split('/')[8])
				} else { $member_vm = $($vm.id.split('/')[8]) }
				
				$count += 1
				if($count -eq 5) {
					$member_vm = $member_vm + "`n"
					$count = 0
				}
			}
			$uregwksht.Cells.Item($row, $col++) = "$member_vm".tolower()
			$datafound=1
		}
	}
		if(-not $datafound){ $row=2;$col=1 ;$uregwksht.Cells.Item($row,$col) = "This subscription do not have any Proximity Placement groups"	}

	
	$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
	$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)

	$workbook.Worksheets['Proximity Placement Group'].UsedRange.Columns.Autofit() | Out-Null


}

Function get_StorageAccount {
	if($global:mode -ne 15) { ascii }
	logger -t i -msg "Collecting Storage Account details..."
	$SAs = Get-AzStorageAccount
	$datafound=0
	#excel creation
	$row = 1; $col = 1
	$uregwksht= $workbook.Worksheets.Add()
	$uregwksht.Name = 'Storage Account'

	if ($SAs) {
		for($i = 1; $i -le 5; $i++) {
			$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
			$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}
		
		$uregwksht.Cells.Item($row,$col++) = "Account Name"
		$uregwksht.Cells.Item($row,$col++) = "Resource Group"
		$uregwksht.Cells.Item($row,$col++) = "Location"
		$uregwksht.Cells.Item($row,$col++) = "SKU"
		$uregwksht.Cells.Item($row,$col++) = "Access Tier"
		
		foreach($sa in $SAs) {
			$row++; $col = 1
			$uregwksht.Cells.Item($row, $col++) = "$($sa.StorageAccountName)".tolower()
			$uregwksht.Cells.Item($row, $col++) = "$($sa.ResourceGroupName)".tolower()
			$uregwksht.Cells.Item($row, $col++) = "$($sa.PrimaryLocation)".tolower()
			$uregwksht.Cells.Item($row, $col++) = "$($sa.sku.name) "
			$uregwksht.Cells.Item($row, $col++) = "$($sa.Accesstier)".tolower()
			$datafound=1
		
		}
	
		$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
		$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)
		$workbook.Worksheets['Storage Account'].UsedRange.Columns.Autofit() | Out-Null

	}	
	
		if(-not $datafound){ $row=2;$col=1 ;$uregwksht.Cells.Item($row,$col) = "This subscription do not have any Storage Accounts"	}
	
	$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
	$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)
	$workbook.Worksheets['Storage Account'].UsedRange.Columns.Autofit() | Out-Null


}

Function get_VirtualNetwork  {
	if($global:mode -ne 15) { ascii }
	logger -t i -msg "Collecting Virtual Network details..."
	$vnet = Get-AzVirtualNetwork
	$datafound=0
	#excel creation
	$row = 1; $col = 1
	$uregwksht= $workbook.Worksheets.Add()
	$uregwksht.Name = 'VNET'

	if($vnet) {
		for($i = 1; $i -le 4; $i++) {
			$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
			$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}
		$uregwksht.Cells.Item($row,$col++) = "Virtual Network"
		$uregwksht.Cells.Item($row,$col++) = "Address Space"
		$uregwksht.Cells.Item($row,$col++) = "Resource Group"
		$uregwksht.Cells.Item($row,$col++) = "Location"

		for ($i=0 ; $i -lt $vnet.Name.count; $i++)
		{
			$row++ ; $col = 1
			$uregwksht.Cells.Item($row, $col++) = "$($vnet[$i].name)".tolower()
			$uregwksht.Cells.Item($row, $col++) = "$($vnet[$i].addressSpace.addressPrefixes)".tolower()
			$uregwksht.Cells.Item($row, $col++) = "$($vnet[$i].resourceGroupname)".tolower()
			$uregwksht.Cells.Item($row, $col++) = "$($vnet[$i].location)".tolower()
			$datafound=1
		}
	}
	
		if(-not $datafound){ $row=2;$col=1 ;$uregwksht.Cells.Item($row,$col) = "This subscription do not have any Virtual Network"	}

	$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
	$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)
	$workbook.Worksheets['VNET'].UsedRange.Columns.Autofit() | Out-Null

#	Start the subnet information sheet

	Function Subnet{
		logger -t i -msg "Collecting Subnet details..."
		#excel creation
		$row = 1; $col = 1
		$uregwksht= $workbook.Worksheets.Add()
		$uregwksht.Name = 'Subnet'
		$datafound=0
		
		for($i = 1; $i -le 5; $i++) 
		{
			$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
			$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}
		$uregwksht.Cells.Item($row,$col++) = "Subnet Name"
		$uregwksht.Cells.Item($row,$col++) = "NsgMap"
		$uregwksht.Cells.Item($row,$col++) = "subnetAddressPrefix"
		$uregwksht.Cells.Item($row,$col++) = "VirtualNetwork"
		$uregwksht.Cells.Item($row,$col++) = "vnetAddressPrefix"

		for ($i=0;$i -lt $vnet.count ;$i++) {
			$subnet = Get-AzVirtualNetworkSubnetConfig -VirtualNetwork $vnet[$i]
			
			if($subnet) {
				for ($j=0;$j -lt $subnet.count; $j++) {
					$row++ ; $col = 1
					$uregwksht.Cells.Item($row, $col++) = "$($subnet[$j].name)".tolower()
					
					if ( $subnet[$j].networkSecurityGroup) {
						$uregwksht.Cells.Item($row, $col++) = "$($subnet[$j].networkSecurityGroup.id.split('/')[8])".tolower()
					}
					else {
						$uregwksht.Cells.Item($row, $col++) = "NA"
					}
					
					$uregwksht.Cells.Item($row, $col++) = "$($subnet[$j].addressPrefix)".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$($vnet[$i].name)".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$($vnet[$i].addressSpace.addressPrefixes)".tolower()
					$datafound=1
				}
			}
		}
	
		if(-not $datafound){ $row=2;$col=1 ;$uregwksht.Cells.Item($row,$col) = "This subscription do not have any Network Subnet"	}

		$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
		$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)

		$workbook.Worksheets['Subnet'].UsedRange.Columns.Autofit() | Out-Null

}
	Subnet
}

Function get_VirtualMachine {
	if($global:mode -ne 15) { ascii }
	logger -t i -msg "Collecting Virtual Machine details..."
	$vm = Get-AzVM
	$datafound=0
	#excel creation
	$row = 1; $col = 1
	$uregwksht= $workbook.Worksheets.Add()
	$uregwksht.Name = 'VM'
	
	if($vm) {
		for($i = 1; $i -le 6; $i++) {
			$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
			$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}
		$uregwksht.Cells.Item($row,$col++) = "VM Name"
		$uregwksht.Cells.Item($row,$col++) = "Resource Group"
		$uregwksht.Cells.Item($row,$col++) = "SKU"
		$uregwksht.Cells.Item($row,$col++) = "Availability Set"
		$uregwksht.Cells.Item($row,$col++) = "PPG Name"
		#$uregwksht.Cells.Item($row,$col++) = "Disk Name"
		#$uregwksht.Cells.Item($row,$col++) = "Disk Size"
		$uregwksht.Cells.Item($row,$col++) = "Extensions"

		$uregwksht.columns("F").columnwidth = 50
		
	for ($i =0; $i -lt $vm.count; $i++) 
	{
		$extensions=""
		for ($j=0;$j -lt $($vm[$i].Extensions.id.count); $j++)
		{
			if ($extensions -ne "") 
			{
				$extensions += "`n"+ $($vm[$i].Extensions.id[$j].split('/')[10])
			} 
			else 
			{
				$extensions += $($vm[$i].Extensions.id[0].split('/')[10])
			}
		}
		
		$row++; $col = 1
		$uregwksht.Cells.Item($row, $col++) = "$($vm[$i].name)".tolower()
		$uregwksht.Cells.Item($row, $col++) = "$($vm[$i].resourceGroupname)".tolower()
		$uregwksht.Cells.Item($row, $col++) = "$($vm[$i].hardwareProfile.vmSize)".tolower()
		if ($($vm[$i].AvailabilitySetReference.id)) {
			$uregwksht.Cells.Item($row, $col++) = "$($vm[$i].AvailabilitySetReference.id.split('/')[8]) "
			if ($($vm[$i].proximityPlacementGroup)) { $uregwksht.Cells.Item($row, $col++) = "$($vm[$i].proximityPlacementGroup.id.split('/')[8] )".tolower()	}
			else {$uregwksht.Cells.Item($row, $col++) = "NA" }
		}
		else {
			$uregwksht.Cells.Item($row, $col++) = "NA"
			$uregwksht.Cells.Item($row, $col++) = "NA"			
		}
		#$uregwksht.Cells.Item($row, $col++) = "$($vm[$i].storageProfile.datadisks.name)".tolower()
		#$uregwksht.Cells.Item($row, $col++) = "$($vm[$i].storageProfile.datadisks.diskSizeGb)".tolower()
		$uregwksht.Cells.Item($row, $col++) = "$extensions"
		$uregwksht.Cells.Item($row,$col++) = "$global:subscription"
		$datafound=1
	}
	

	
	$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
	$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)
	$workbook.Worksheets['VM'].UsedRange.Columns.Autofit() | Out-Null
	

	}
		

}

Function get_NetworkInterface {
	if($global:mode -ne 15) { ascii }
	logger -t i -msg "Collecting Network Interface details..."
	$nic = Get-AzNetworkInterface
	$datafound=0
	#excel creation
	$row = 1; $col = 1
	$uregwksht= $workbook.Worksheets.Add()
	$uregwksht.Name = 'NICforVM'
	
	if($nic){
		for($i = 1; $i -le 11; $i++) {
		$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
		$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}
		$uregwksht.Cells.Item($row,$col++) = "VM Name"
		$uregwksht.Cells.Item($row,$col++) = "NIC name"
		$uregwksht.Cells.Item($row,$col++) = "VNET Name"
		$uregwksht.Cells.Item($row,$col++) = "VNET Resource Group"
		$uregwksht.Cells.Item($row,$col++) = "Subnet"
		$uregwksht.Cells.Item($row,$col++) = "PrivateIPAllocationMethod"
		$uregwksht.Cells.Item($row,$col++) = "IP Address"
		$uregwksht.Cells.Item($row,$col++) = "Public IP"
		$uregwksht.Cells.Item($row,$col++) = "IP Forwarding"
		$uregwksht.Cells.Item($row,$col++) = "NIC Resource Group"
		$uregwksht.Cells.Item($row,$col++) = "Accelerated Networking"
		
		

		for ($i = 0; $i -lt $nic.count; $i++){
			$start_row = $row+1
			if($($nic[$i].ipConfigurations.count) -ge 1) {
				for($j=0;$j -lt $nic[$i].ipConfigurations.count ; $j++) {
					$row++; $col = 1
					if (-not $($nic[$i].VirtualMachine.id) ){
						$uregwksht.Cells.Item($row,$col).EntireRow.Font.ColorIndex = 3  
						$uregwksht.Cells.Item($row, $col++) = "Not Attached"
					}
					else{
						$uregwksht.Cells.Item($row, $col++) = "$($nic[$i].virtualMachine.id.split('/')[8])".tolower()
					}
					$uregwksht.Cells.Item($row, $col++) = "$($nic[$i].name)".tolower()

					$uregwksht.Cells.Item($row, $col++) = "$($nic[$i].ipConfigurations[$j].subnet.id.split('/')[8])".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$($nic[$i].ipConfigurations[$j].subnet.id.split('/')[4])".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$($nic[$i].ipConfigurations[$j].subnet.id.split('/')[10])".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$($nic[$i].ipConfigurations[$j].privateIpAllocationMethod)".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$($nic[$i].ipConfigurations[$j].privateIpAddress)".tolower()
					$datafound=1
					if ($($nic[$i].ipConfigurations[$j].publicIpAddress)) {
							$uregwksht.Cells.Item($row, $col++) = "$($nic[$i].ipConfigurations[$j].publicIpAddress)".tolower()
						}
					else {
							$uregwksht.Cells.Item($row, $col++) = "No"
					}
					$uregwksht.Cells.Item($row, $col++) = "$($nic[$i].enableIpForwarding)".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$($nic[$i].resourceGroupname)".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$($nic[$i].EnableAcceleratedNetworking)".tolower()
				}
			}
			$end_row = $row
			
			function merge_cell ($col, $srow, $erow) 
			{
			$MergeCells = $uregwksht.Range("$col${srow}:$col$erow")
			$MergeCells.Select() |Out-Null
			$MergeCells.MergeCells = $true  
			}
			
			if ($($nic[$i].ipConfigurations.count) -gt 1) {
					merge_cell "A" $start_row $end_row
					merge_cell "B" $start_row $end_row
					merge_cell "C" $start_row $end_row
					merge_cell "D" $start_row $end_row
					merge_cell "E" $start_row $end_row
					merge_cell "H" $start_row $end_row
					merge_cell "I" $start_row $end_row
					merge_cell "J" $start_row $end_row
					merge_cell "K" $start_row $end_row
				}

		}
	}
		if(-not $datafound){ $row=2;$col=1 ;$uregwksht.Cells.Item($row,$col) = "This subscription do not have any Network Interfaces"	}

	$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
	$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)

	$workbook.Worksheets['NICforVM'].UsedRange.Columns.Autofit() | Out-Null

}

Function get_Netapp { 
	if($global:mode -ne 15) { ascii }
	logger -t i -msg "Collecting Azure Netapp Files details..."
	$rg = Get-AzResourceGroup
	$dadtafound=0
	
		#excel creation
		$row = 1; $col = 1
		$uregwksht= $workbook.Worksheets.Add()
		$uregwksht.Name = 'ANF Summary'
		
		for($i = 1; $i -le 8; $i++) {
		$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
		$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}
		$uregwksht.Cells.Item($row,$col++) = "ANF Resource Group"
		$uregwksht.Cells.Item($row,$col++) = "ANF Account Name"
		$uregwksht.Cells.Item($row,$col++) = "ANF Capacity Pool"
		$uregwksht.Cells.Item($row,$col++) = "ANF Volume"
		$uregwksht.Cells.Item($row,$col++) = "Volume Size (GB)".tolower()
		$uregwksht.Cells.Item($row,$col++) = "Tier"
		$uregwksht.Cells.Item($row,$col++) = "Subnet"
		

	for ($i=0; $i -lt $rg.count; $i++) { 
		$anfaccount = Get-AzNetappFilesaccount -ResourceGroupName "$($rg[$i].ResourceGroupname)".tolower()
		
		if($anfaccount) {
			for($j=0;$j -lt $anfaccount.count ; $j++) {
				$anfcappool = Get-AzNetAppFilesPool -ResourceGroupName  $($anfaccount[$j].ResourceGroupName) -AccountName $($anfaccount[$j].Name)
				
				if ($anfcappool) {
					for ($k=0;$k -lt $anfcappool.count ; $k++){
						$anfvolume = Get-AzNetAppFilesVolume -ResourceGroupName $($anfaccount[$j].ResourceGroupName)  -AccountName $($anfaccount[$j].Name) -PoolName  $($anfcappool[$k].name.split('/')[1])
						 
						if($anfvolume) {
							for($m=0;$m -lt $anfvolume.count ; $m++) {
								$row++ ; $col=1
								
								$uregwksht.Cells.Item($row,$col++) = "$($anfaccount[$j].ResourceGroupName)".tolower()
								$uregwksht.Cells.Item($row,$col++) = "$($anfaccount[$j].Name) "
								$uregwksht.Cells.Item($row,$col++) = "$($anfcappool[$k].name.split('/')[1])".tolower()
								$uregwksht.Cells.Item($row,$col++) = "$($anfvolume[$m].Name.split('/')[2])".tolower()
								$uregwksht.Cells.Item($row,$col++) = "$($anfvolume[$m].UsageThreshold/1024/1024/1024) GB"
								$uregwksht.Cells.Item($row,$col++) = "$($anfvolume[$m].ServiceLevel)".tolower()
								$uregwksht.Cells.Item($row,$col++) = "$($anfvolume[$m].Subnetid.split('/')[10])".tolower()
								$datafound=1
							}
						}
					}
				}
			}
		}
		
	}
	if(-not $datafound){ $row=2;$col=1 ;$uregwksht.Cells.Item($row,$col) = "This subscription do not have any Azure Netapp Files"	}

	
	$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
	$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)

	$workbook.Worksheets['ANF Summary'].UsedRange.Columns.Autofit() | Out-Null

}

Function get_Disk {
	if($global:mode -ne 15) { ascii }
	logger -t i -msg "Collecting Managed Disk details..."
	$disk = Get-AzDisk
	$datafound=0
	#excel creation
	$row = 1; $col = 1
	$uregwksht= $workbook.Worksheets.Add()
	$uregwksht.Name = 'DiskforVM'
	
		for($i = 1; $i -le 6; $i++) {
		$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
		$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}
		$uregwksht.Cells.Item($row,$col++) = "VM Name"
		$uregwksht.Cells.Item($row,$col++) = "Disk Name"
		$uregwksht.Cells.Item($row,$col++) = "Disk Size"
		$uregwksht.Cells.Item($row,$col++) = "SKU"
		$uregwksht.Cells.Item($row,$col++) = "Tier"
		$uregwksht.Cells.Item($row,$col++) = "Disk State"
		
		if($disk) {
			for ($i = 0 ; $i -lt $disk.count; $i++) 
			{	
				$row++; $col=1
				if($($disk[$i].managedby)) {	$uregwksht.Cells.Item($row,$col++) = "$($disk[$i].managedby.split('/')[8])".tolower() }
				else {  $uregwksht.Cells.Item($row,$col++) = "Disk Not Attached";  $uregwksht.Cells.Item($row,$col).EntireRow.Font.ColorIndex = 3  }
				$uregwksht.Cells.Item($row,$col++) = "$($disk[$i].name)".tolower()
				$uregwksht.Cells.Item($row,$col++) = "$($disk[$i].diskSizeGb) GB"
				$uregwksht.Cells.Item($row,$col++) = "$($disk[$i].sku.name)".tolower()
				$uregwksht.Cells.Item($row,$col++) = "$($disk[$i].sku.Tier)".tolower()
				$uregwksht.Cells.Item($row,$col++) = "$($disk[$i].diskState)".tolower()
				$datafound=1
			}
		}
		if(-not $datafound){ $row=2;$col=1 ;$uregwksht.Cells.Item($row,$col) = "This subscription do not have any Virtual Network"	}
		
			$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
			$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)
	
			$workbook.Worksheets['DiskforVM'].UsedRange.Columns.Autofit() | Out-Null
	


}

Function get_LoadBalancer {

	Function LB_BackendPool
	{
		logger -t i -msg "Collecting LB Backend details..."
		#excel creation
		$row = 1; $col = 1
		$uregwksht= $workbook.Worksheets.Add()
		$uregwksht.Name = 'LB Backend Pool'
			
		for($i = 1; $i -le 5; $i++) 
		{
			$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
			$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}
		
		$uregwksht.Cells.Item($row,$col++) = "ResourceGroupName"
		$uregwksht.Cells.Item($row,$col++) = "BackEnd Pool Name"
		$uregwksht.Cells.Item($row,$col++) = "LB Name"
		$uregwksht.Cells.Item($row,$col++) = "VNET/Subnet Name"
		$uregwksht.Cells.Item($row,$col++) = "Member VM Name"
		
		for ($m=0; $m -lt $LB.count; $m++) {
			for ($i = 0; $i -lt $($LB[$m].backendAddressPools.count); $i++)
			{
				$row++; $col = 1
				$uregwksht.Cells.Item($row, $col++) = "$($LB[$m].resourceGroupname)".tolower()
				$uregwksht.Cells.Item($row, $col++) = "$($LB[$m].backendAddressPools[$i].name)".tolower()
				$uregwksht.Cells.Item($row, $col++) = "$($LB[$m].name)".tolower()
				$vnetsubnetname= "$($LB[$m].frontendIpConfigurations[0].subnet.id.split('/')[8])".tolower()+"/"+"$($LB[$m].frontendIpConfigurations[0].subnet.id.split('/')[10])".tolower()
				$uregwksht.Cells.Item($row, $col++) = "$vnetsubnetname"
				
				if ($($LB[$m].BackendAddressPools[$i].BackendIpConfigurations)) {
					$NetConfCount = $LB[$m].BackendAddressPools[$i].BackendIpConfigurations.count
					$vmname = ""
					if ($NetConfCount -eq 1) { 
						$length = $LB[$m].BackendAddressPools[$i].BackendIpConfigurations.id.split('/')[8].length 
						$vmname =  $LB[$m].BackendAddressPools[$i].BackendIpConfigurations.id.split('/')[8].substring(0,$length-6)
					}else {
						for ($j = 0; $j -lt $NetConfCount ; $j++)
						{
							$length = $LB[$m].BackendAddressPools[$i].BackendIpConfigurations.id[$j].split('/')[8].length 
							if ($vmname -eq "" ) { 	$vmname +=  $LB[$m].BackendAddressPools[$i].BackendIpConfigurations.id[$j].split('/')[8].substring(0,$length-6) 	}
							else {	$vmname += "," + $LB[$m].BackendAddressPools[$i].BackendIpConfigurations.id[$j].split('/')[8].substring(0,$length-6) }
						}	
					}
					$uregwksht.Cells.Item($row, $col++) = "$vmname"
				}
				else {	$uregwksht.Cells.Item($row, $col++) = "No members" }
			}
		}
		
			$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
			$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)
	
			$workbook.Worksheets['LB BackEnd Pool'].UsedRange.Columns.Autofit() | Out-Null
	
		#LOGGER "Gather Load Balancer BackEnd Pool details"  "Completed"
	
	
	
	}
		
	function LB_Frontend
	{
		logger -t i -msg "Collecting LB Frontend details..."
		
		#excel creation
		$row = 1; $col = 1
		$uregwksht= $workbook.Worksheets.Add()
		$uregwksht.Name = 'LBFrontendIP'
			
		for($i = 1; $i -le 5; $i++) 
		{
			$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
			$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}
			
		$uregwksht.Cells.Item($row,$col++) = "ResourceGroupName"
		$uregwksht.Cells.Item($row,$col++) = "LB Frontend IP Name"
		$uregwksht.Cells.Item($row,$col++) = "LB Name"
		$uregwksht.Cells.Item($row,$col++) = "VNET/Subnet Name"
		$uregwksht.Cells.Item($row,$col++) = "IP Address"
		
		for ($m=0; $m -lt $LB.count; $m++) {
			if ($LB[$m].frontendIpConfigurations) {
				for ($i = 0; $i -lt $($LB[$m].frontendIpConfigurations.count); $i++){
					$row++; $col = 1
					$uregwksht.Cells.Item($row, $col++) = "$($LB[$m].resourceGroupname)".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$($LB[$m].frontendIpConfigurations[$i].name)".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$($LB[$m].name)".tolower()
					$vnetsubnetname= "$($LB[$m].frontendIpConfigurations[$i].subnet.id.split('/')[8])".tolower()+"/"+"$($LB[$m].frontendIpConfigurations[$i].subnet.id.split('/')[10])".tolower()
					$uregwksht.Cells.Item($row, $col++) = "$vnetsubnetname"					
					$uregwksht.Cells.Item($row, $col++) = "$($LB[$m].frontendIpConfigurations[$i].privateIpAddress)".tolower()
					}
				}
			}
		
		$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
		$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)
	
		$workbook.Worksheets['LBFrontendIP'].UsedRange.Columns.Autofit() | Out-Null
	
	}

		
	function LB_Probes {
		
		logger -t i -msg "Collecting LB Probes Details..."
		
		#excel creation
		$row = 1; $col = 1
		$uregwksht= $workbook.Worksheets.Add()
		$uregwksht.Name = 'LBprobes'
		
		for($i = 1; $i -le 7; $i++) {
			$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
			$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
			}
		
		$uregwksht.Cells.Item($row,$col++) = "ResourceGroupName"
		$uregwksht.Cells.Item($row,$col++) = "Probe Name"
		$uregwksht.Cells.Item($row,$col++) = "LB Name"
		$uregwksht.Cells.Item($row,$col++) = "Protocol"
		$uregwksht.Cells.Item($row,$col++) = "Port"
		$uregwksht.Cells.Item($row,$col++) = "Interval"
		$uregwksht.Cells.Item($row,$col++) = "NumberOfProbes"
	
		for ($m=0; $m -lt $LB.count;$m++) {
			if ($($LB[$m].probes)) {
				for ($i=0;$i -lt $($LB[$m].probes.count) ; $i++) {
					$row++; $col=1
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].resourceGroupname)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].probes[$i].name)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].name)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].probes[$i].protocol)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].probes[$i].port)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].probes[$i].IntervalInSeconds)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].probes[$i].NumberOfProbes)".tolower()
				}
			}
		}
		
		$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
		$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)
	
		$workbook.Worksheets['LBprobes'].UsedRange.Columns.Autofit() | Out-Null
	
	
	}


	function LB_Rules {
		
		logger -t i -msg "Collecting LB rules details..."
		#excel creation
		$row = 1; $col = 1
		$uregwksht= $workbook.Worksheets.Add()
		$uregwksht.Name = 'LBrules'
		
		for($i = 1; $i -le 8; $i++) 
		{
			$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
			$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}
		
		$uregwksht.Cells.Item($row,$col++) = "ResourceGroupName"
		$uregwksht.Cells.Item($row,$col++) = "Rule Name"
		$uregwksht.Cells.Item($row,$col++) = "BackEnd Name"
		$uregwksht.Cells.Item($row,$col++) = "Health Probe"
		$uregwksht.Cells.Item($row,$col++) = "TCP reset"
		$uregwksht.Cells.Item($row,$col++) = "IdleTimeOut"
		$uregwksht.Cells.Item($row,$col++) = "Enable FloatingIP"
		$uregwksht.Cells.Item($row,$col++) = "FrontEnd IP"
		
		for ($m=0; $m -lt $LB.count ; $m++) {
			if ($($LB[$m].LoadBalancingRules)) {
				for ($i=0; $i -lt $($LB[$m].LoadBalancingRules.count);$i++) {
					$row++; $col=1
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].resourceGroupname)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].LoadBalancingRules[$i].name)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].LoadBalancingRules[$i].BackendAddressPool.id.split('/')[10])".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].LoadBalancingRules[$i].Probe.id.split('/')[10])".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].LoadBalancingRules[$i].EnableTcpReset)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].LoadBalancingRules[$i].IdleTimeoutInMinutes)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].LoadBalancingRules[$i].EnableFloatingIP)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($LB[$m].LoadBalancingRules[$i].FrontendIPConfiguration.id.split('/')[10])".tolower()
				}
			}
		}
		
		$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
		$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)
	
		$workbook.Worksheets['LBrules'].UsedRange.Columns.Autofit() | Out-Null
	
	
	}
	



	if($global:mode -ne 15) { ascii }
	logger -t i -msg "Collecting Load Balancer details..."
	$LB = Get-AzLoadBalancer
	$datafound=0
	#excel creation
	$row = 1; $col = 1
	$uregwksht= $workbook.Worksheets.Add()
	$uregwksht.Name = 'Load Balancer'

	if($LB) {
		for($i = 1; $i -le 3; $i++) {
			$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
			$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
			}
		$uregwksht.Cells.Item($row,$col++) = "ResourceGroupName"
		$uregwksht.Cells.Item($row,$col++) = "LB Name"
		$uregwksht.Cells.Item($row,$col++) = "Subnet Mapping"
		
		for ($i = 0; $i -lt $LB.count; $i++){
			$row++; $col = 1
			$uregwksht.Cells.Item($row, $col++) = "$($LB[$i].resourceGroupname)".tolower()
			$uregwksht.Cells.Item($row, $col++) = "$($LB[$i].name)".tolower()
			$uregwksht.Cells.Item($row, $col++) = "$($LB[$i].frontendIpConfigurations[0].subnet.id.split('/')[10])".tolower()
			$datafound=1
			}
		
		$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
		$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)
	
		$workbook.Worksheets['Load Balancer'].UsedRange.Columns.Autofit() | Out-Null
	
		#calling the sub-functions	
		LB_BackendPool
		LB_Probes
		LB_Rules
		LB_Frontend

		}
		
		if(-not $datafound){ $row=2;$col=1 ;$uregwksht.Cells.Item($row,$col) = "This subscription do not have any Load Balancer"	}
	
}

Function get_Routes {
	if($global:mode -ne 15) { ascii }

}

Function get_NetworkSecurityGroups {
	if($global:mode -ne 15) { ascii }
	logger -t i -msg "Collecting Network Security Group details..."
	$nsgs = Get-AzNetworkSecurityGroup
	$datafound=0
		#excel creation
		$row = 1; $col = 1
		$uregwksht= $workbook.Worksheets.Add()
		$uregwksht.Name = 'NetworkSecurityGroup'
	
	
	for($i = 1; $i -le 10; $i++) {
		$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
		$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}

		$uregwksht.Cells.Item($row,$col++) = "NSG Name"
		$uregwksht.Cells.Item($row,$col++) = "Direction"
		$uregwksht.Cells.Item($row,$col++) = "Priority"
		$uregwksht.Cells.Item($row,$col++) = "Rule Name"
		$uregwksht.Cells.Item($row,$col++) = "Source Port Range"
		$uregwksht.Cells.Item($row,$col++) = "Destination Port Range"
		$uregwksht.Cells.Item($row,$col++) = "Protocol"
		$uregwksht.Cells.Item($row,$col++) = "SourceAddressPrefix"
		$uregwksht.Cells.Item($row,$col++) = "DestinationAddressPrefix"
		$uregwksht.Cells.Item($row,$col++) = "Access"
		

	if($nsgs) {
		
		foreach ($nsg in $nsgs) {
			if($nsg.securityRules) {
				foreach($rule in $($nsg.SecurityRules)) {
						$row++ ; $col=1
						$uregwksht.Cells.Item($row,$col++) = "$($nsg.name)".tolower()
						$uregwksht.Cells.Item($row,$col++) = "$($rule.direction)".tolower()
						$uregwksht.Cells.Item($row,$col++) = "$($rule.Priority)".tolower()
						$uregwksht.Cells.Item($row,$col++) = "$($rule.name)".tolower()
						$datafound=1
						if ($($rule.SourcePortRange) -eq "*") 
							{ 
								$sourceportrange="ANY" 
							}
						else 
							{ 
								$sourceportrange="$($rule.SourcePortRange)".tolower() 
							}
						
						$uregwksht.Cells.Item($row,$col++) = "$sourceportrange"
						
						if ($($rule.DestinationPortRange) -eq "*") 
							{
								$DestinationPortRange="ANY" 
							}
						else 
							{ 
								$DestinationPortRange="$($rule.DestinationPortRange)".tolower()
							}
						
						$uregwksht.Cells.Item($row,$col++) = "$DestinationPortRange"
	
						if ($($rule.Protocol) -eq "*") 
							{ 
								$Protocol="ANY" 
							}
						else 
							{ 
								$Protocol="$($rule.Protocol)".tolower() 
							}
	
						$uregwksht.Cells.Item($row,$col++) = "$Protocol"
	
						if ($($rule.SourceAddressPrefix) -eq "*") 
							{ 
								$SourceAddressPrefix="ANY" 
							}
						else 
							{ 
								$SourceAddressPrefix="$($rule.SourceAddressPrefix)".tolower() 
							}
						
						$uregwksht.Cells.Item($row,$col++) = "$SourceAddressPrefix"
						
						if ($($rule.DestinationAddressPrefix) -eq "*") 
							{ 
								$DestinationAddressPrefix="ANY" 
							}
						else 
							{ 
								$DestinationAddressPrefix="$($rule.DestinationAddressPrefix)".tolower()
							}
						
						$uregwksht.Cells.Item($row,$col++) = "$DestinationAddressPrefix"
						
						if($($rule.Access) -eq "Deny")
							{ 
								$uregwksht.Cells.Item($row,$col).EntireRow.Font.ColorIndex = 3  
							}
						$uregwksht.Cells.Item($row,$col++) = "$($rule.Access)".tolower()
					
				}
			}
			
			foreach ($rule in $($nsg.DefaultSecurityRules)) {
					$row++ ; $col=1
					$uregwksht.Cells.Item($row,$col++) = "$($nsg.name)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($rule.direction)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($rule.Priority)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($rule.name)".tolower()
					$datafound=1
					if ($($rule.SourcePortRange) -eq "*") { $sourceportrange="ANY" }
					else { $sourceportrange="$($rule.SourcePortRange)".tolower() }
					
					$uregwksht.Cells.Item($row,$col++) = "$sourceportrange"
					
					if ($($rule.DestinationPortRange) -eq "*") { $destportrange="ANY" }
					else { $destportrange="$($rule.DestinationPortRange)".tolower() }
					
					$uregwksht.Cells.Item($row,$col++) = "$destportrange"
					
					if ($($rule.Protocol) -eq "*") { $Protocol="ANY" }
					else { $Protocol="$($rule.Protocol)".tolower() }
				
					$uregwksht.Cells.Item($row,$col++) = "$Protocol"
					
					if ($($rule.SourceAddressPrefix) -eq "*") { $SourceAddressPrefix="ANY" }
					else { $SourceAddressPrefix="$($rule.SourceAddressPrefix)".tolower() }
				
					$uregwksht.Cells.Item($row,$col++) = "$SourceAddressPrefix"
					
					if ($($rule.DestinationAddressPrefix) -eq "*") { $DestinationAddressPrefix="ANY" }
					else { $DestinationAddressPrefix="$($rule.DestinationAddressPrefix)".tolower() }
					
					$uregwksht.Cells.Item($row,$col++) = "$DestinationAddressPrefix"
					
					if ($($rule.Access) -eq "Deny") { $uregwksht.Cells.Item($row,$col).EntireRow.Font.ColorIndex = 3 }

					$uregwksht.Cells.Item($row,$col++) = "$($rule.Access)".tolower()
				}
			
		}
	}		


	if(-not $datafound){ $row=2;$col=1 ;$uregwksht.Cells.Item($row,$col) = "This subscription do not have any Network Security Groups"	}
	
	$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
	$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)
	$workbook.Worksheets['NetworkSecurityGroup'].UsedRange.Columns.Autofit() | Out-Null

}

Function get_ApplicationSecurityGroups {
	if($global:mode -ne 15) { ascii }
	logger -t i -msg "Collecting Application Security Group details.."
	$datafound=0
	#$asg = Get-AzApplicationSecurityGroup
	$asgs = Get-AzNetworkInterface
		
			#excel creation
		$row = 1; $col = 1
		$uregwksht= $workbook.Worksheets.Add()
		$uregwksht.Name = 'ApplicationSecurityGroup'
		
	for($i = 1; $i -le 3; $i++) {
		$uregwksht.Cells.item(1,$i).EntireRow.Font.Bold = $True ;
		$uregwksht.Cells.item(1,$i).Interior.ColorIndex = 44
		}

		$uregwksht.Cells.Item($row,$col++) = "ASG Name"
		$uregwksht.Cells.Item($row,$col++) = "Resource Group"
		$uregwksht.Cells.Item($row,$col++) = "NICs"

	if($asgs) {
		foreach($nic in $asgs) {
			foreach($ip in $($nic.ipConfigurations)) {
				if($($ip.ApplicationSecurityGroups)) {
					$row++ ; $col=1
					$datafound=1
					$uregwksht.Cells.Item($row,$col++) = "$($ip.ApplicationSecurityGroups.id.split('/')[8])".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($nic.resourceGroupname)".tolower()
					$uregwksht.Cells.Item($row,$col++) = "$($nic.name)".tolower()
				}
			}
		}
	}
	
		if(-not $datafound){ $row=2;$col=1 ;$uregwksht.Cells.Item($row,$col) = "This subscription do not have any Application security groups"	}
	
	$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
	$uregwksht.move([System.Reflection.Missing]::Value, $lastSheet)
	$workbook.Worksheets['ApplicationSecurityGroup'].UsedRange.Columns.Autofit() | Out-Null
	
}

Function get_Firewall {
	if($global:mode -ne 15) { ascii }

	$firewall = Get-AzFirewall
	
}


Function subscriptionMenu {
	clear-host
	ascii
	get_SubscriptionList
	Write-Host "`n`t`t0. Exit"
	Write-Host "`t--------------------------------------------------------------"
	[int]$global:sel_subs = Read-Host "`n`tSelect Subscription"
	
	if ($($global:sel_subs) -gt $($global:subs.count)) 
		{
			Write-host "incorrect option selected. Exit" -ForegroundColor Red
			exit
		}
	
	if ($($global:sel_subs) -eq 0 ) 
		{ 
			Write-Host "`nGood Bye ..`n"
			exit 
		}

	logger -t i -msg "Setting Azure context to subscription $($global:subs.Subscription.Name[$($global:sel_subs -1)])".tolower()
	set-AzContext -Subscription $($global:subs.Subscription.Name[$($global:sel_subs - 1)]) |Out-null
	
	resourceMenu
}


Function resourceMenu {
	clear-host
	ascii

	Write-Output "`n`t`t Select any individual / all component : `n"
	Write-Host "`t`t" "1.".padleft(4) " Resource Group"
	Write-Host "`t`t" "2.".padleft(4) " Availability Set"
	Write-Host "`t`t" "3.".padleft(4) " Proximity Placement Group"
	Write-Host "`t`t" "4.".padleft(4) " Storage Account"
	Write-Host "`t`t" "5.".padleft(4) " Virtual Network"
	Write-Host "`t`t" "6.".padleft(4) " Virtual Machine"
	Write-Host "`t`t" "7.".padleft(4) " Network Interface"
	Write-Host "`t`t" "8.".padleft(4) " Azure Netapp Files (ANF)".tolower()
	Write-Host "`t`t" "9.".padleft(4) " Disks"
	Write-Host "`t`t" "10.".padleft(4) " Load Balancers"
	Write-Host "`t`t" "11.".padleft(4) " Route Tables (UDR)".tolower()
	Write-Host "`t`t" "12.".padleft(4) " Network Security Group (NSG)".tolower()
	Write-Host "`t`t" "13.".padleft(4) " Application Security Group (ASG)".tolower()
	Write-Host "`t`t" "14.".padleft(4) " Firewall"
	Write-Host "`n`t`t" "15.".padleft(4) " All Components"
	Write-Host "`n`t`t" "0.".padleft(4) " Exit " 
	Write-Host "`t--------------------------------------------------------------"
	[int]$global:mode = Read-Host "`n`t`t Select option "
	
	ascii
	
}

Function all {
		get_ResourceGroup
		get_AvailabilitySet
		get_PPG
		get_StorageAccount
		get_VirtualNetwork
		get_VirtualMachine
		get_NetworkInterface
		get_Netapp
		get_Disk		
		get_LoadBalancer
		get_Routes
		get_NetworkSecurityGroups
		get_ApplicationSecurityGroups
		get_Firewall
}



#Printing Version details - 
logger -t i -msg "Azure Data Collector $version"

#check if Az module is installed
logger -t i -msg "Checking Pre-requisites"
$module = Get-InstalledModule -Name Az.Accounts -MinimumVersion 2.7.6 -ErrorAction SilentlyContinue

if (! $module)
{
	logger -t e -msg "Module Az is not installed." -ForegroundColor Red
	logger -t e -msg "This script requires Az module at version not lower than 6.5.0 to run."
	logger -t e -msg "Download and install latest version of Az module following link - https://www.powershellgallery.com/packages/Az `n"
	exit
}


# Check if user is logged in to Azure
if ( ! $(Get-AzContext) )
{
	logger -t e -msg "Failed to connect to Azure. Initiating Azure login..." -ForegroundColor  Red
	Connect-AzAccount
	if (! $?) { exit }
}


#Display Menu
subscriptionMenu

#set mode of run
switch ($global:mode)
{
	1 { $mode = "RG" } ; 2 { $mode = "AvSet" } ; 	3 { $mode = "PPG" } ; 	4 { $mode = "StorageAcc" } ;
	5 { $mode = "Vnet" } ; 	6 { $mode = "VM" } ; 	7 { $mode = "NIC" } ; 	8 { $mode = "ANF" } ;
	9 { $mode = "Disk" } ; 	10 { $mode = "LB" } ; 	11 { $mode = "UDR" } ; 	12 { $mode = "NSG" } ; 	
	13 { $mode = "ASG" } ; 	14 { $mode = "Firewall" } ;15 { $mode = "FULL" }
}

#if user reached here that mean resource to  capture is selected, Start excel now in background

#Set log directory
$logdir = "$($PSScriptRoot)\ADC_reports"

# check if log directory exist else create it.
if (Test-Path -Path $($logdir) ) {	$elog = "$($logdir)\$($global:subscription)_output.xlsx".replace('/','_') }
else { md $logdir |out-null }

#Start excel creation

$tdate = get-date -Format "dd_MM_yyyy"
$elog = "$($logdir)\$($global:subs.Subscription.Name[$($global:sel_subs -1)])_output_$($mode)_$($tdate).xlsx".replace('/','_')
$excel = New-Object -ComObject excel.application
$excel.visible = $False
$excel.DisplayAlerts = $False
$workbook = $excel.Workbooks.Add()



switch ($global:mode)
{
	1 { get_ResourceGroup } ; 	2 { get_AvailabilitySet } ; 3 { get_PPG } ; 4 { get_StorageAccount } ;	5 { get_VirtualNetwork } ;
	6 { get_VirtualMachine } ; 	7 { get_NetworkInterface } ; 8 { get_Netapp } ;	9 { get_Disk } ; 	10 { get_LoadBalancer } ; 
	11 { get_Routes } ; 12 { get_NetworkSecurityGroups } ; 	13 { get_ApplicationSecurityGroups } ; 	14 { get_Firewall } ; 	15 { all } ; 
	0 { Write-Host "`nGood Bye .. `n"
		exit 
		}
	Default {
		"Incorrect Option. Please select the right option."
		exit
	}
}


#close and save the excel
if ( $global:mode -eq 15 ) { $workbook.Worksheets("ResourceGroup").Activate() }
else { $workbook.WorkSheets.item("Sheet1").delete() }

$workbook.SaveAs($elog)  | Out-Null
$workbook.Close|Out-Null
$excel.Quit()|Out-Null

$global:mode=0
ascii
logger -t i -msg "Report saved at : $elog" |Out-String -width 30
Read-Host -Prompt "`n Press any key to exit..."
