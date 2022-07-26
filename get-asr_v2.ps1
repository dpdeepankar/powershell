

param (
[Parameter(Mandatory=$false)][switch]$full,
[Parameter(Mandatory=$false)][switch]$grid,
[Parameter(Mandatory=$false,HelpMessage='Enter the virtual machine name')][string]$vm,
[Parameter(Mandatory=$false,HelpMessage='Enter the Recovery service vault name')][string]$rsv,
[Parameter(Mandatory=$false,HelpMessage='Enter the subscription name')][string]$subs,
[Parameter(Mandatory=$false)][switch]$help,
[Parameter(Mandatory=$false)][switch]$h

)


#Variable definitions

if(($help) -or ($h)) {
	Write-Host "Help doc"
	exit
}

Function Header {
	clear-host
	Write-Host ">>>>>>>>> " -Nonewline ; Write-Host "Azure Site Recovery" -ForegroundColor Green -Nonewline; Write-Host "  <<<<<<<<<`n" 
}


Function getRSVComponents {
	
	Write-Host "`nGetting list of protected items"
	#Get the ASR Fabric info
	$fabric = Get-AzRecoveryServicesAsrFabric 
	
	if (-not $fabric) { 
		Write-Host "`nThis vault is not used for Azure Site Recovery. Please select the correct vault and try again`n" -BackgroundColor Red
		exit
	}
	
	foreach ($f in $fabric) {
		# Get the ASR protection container 
		$container = Get-AzRecoveryServicesAsrProtectionContainer -Fabric $f
		
		$protecteditem = Get-AzRecoveryServicesAsrReplicationProtectedItem -ProtectionContainer $container
		
		if($vm) {  vmBasedOutput($protecteditem) }
		else { regularoutput($protecteditem) }
	}
}


Function recovery_vault_menu {
	Header
	
	Write-host "Looking for All Recovery services vault in subscription " -nonewline ; write-host "$global:subscriptionName" -ForegroundColor Green
	Write-Host "This may take a while, please wait..."
	$RecoveryVaultList = Get-AzRecoveryServicesVault
	
	if(-not $RecoveryVaultList) {
		Write-Host "There are no recovery service vaults in $global:subscriptionName subscription" -BackgroundColor Red
		exit 1
	}
	Header
	Write-Host "List of Available Recovery Services Vault: `n"
	$count=1
	foreach($r in $RecoveryVaultList) {
		Write-Host "${count}. $($r.Name)"
		$count++
	}
	Write-Host "`n0. Exit"
	Write-Host "`--------------------------------------------------------------"
	[int]$sel_subs = Read-Host "`nSelect Recovery Vault"
	
	if ($sel_subs -gt $RecoveryVaultList.count) 
		{
			Write-host "incorrect option selected. Exit" -ForegroundColor Red
			exit
		}
	
	if ($sel_subs -eq 0 ) 
		{ 
			Write-Host "`nGood Bye ..`n"
			exit 
		}


	Header ;Write-Host "Set ASR context to " -Nonewline ; Write-Host "$($RecoveryVaultList[$($sel_subs - 1)].name)" -ForegroundColor Green
	$global:rsv1 = Get-AzRecoveryServicesVault -Name $($RecoveryVaultList[$($sel_subs - 1)].name)
	Set-AzRecoveryServicesAsrVaultContext -Vault $global:rsv1 |out-null

	getRSVComponents
}


Function get_SubscriptionList {
	Header
	Write-host "List of Available Subscriptions: `n"
	$global:subs = Get-Azcontext -listAvailable
	
	$count=0
	foreach ($name in $global:subs.Subscription.Name)
		{
			$count++
			Write-host "${count}. $name"
		}


}


Function subscriptionMenu {
	get_SubscriptionList
	Write-Host "`n0. Exit"
	Write-Host "`--------------------------------------------------------------"
	[int]$sel_subs = Read-Host "`nSelect Subscription"
	
	if ($sel_subs -gt $($global:subs.count)) 
		{
			Write-host "incorrect option selected. Exit" -ForegroundColor Red
			exit
		}
	
	if ($sel_subs -eq 0 ) 
		{ 
			Write-Host "`nGood Bye ..`n"
			exit 
		}

	
	
	$global:subscriptionName = "$($global:subs.Subscription.Name[$($sel_subs - 1)])"
	Write-host "`nSetting Azure context to subscription " -nonewline ; Write-Host "$global:subscriptionName" -ForegroundColor Green
	set-AzContext -Subscription $global:subscriptionName |Out-null
	Header; Write-Host "Set ASR context to " -Nonewline ; Write-Host "$rsv" -ForegroundColor Green		

	if ($rsv) { 
		$recoveryvault = Get-AzRecoveryServicesVault -Name $rsv 
		
		if(-not $recoveryvault) { 
			Write-Host "`nRecovery Vault $rsv does not exist in $global:subscriptionName subscription. Please check the name and try again" -BackgroundColor Red
			exit 1
		}
	
		Set-AzRecoveryServicesAsrVaultContext -Vault $recoveryvault |out-null
		
		getRSVComponents
	}
	else {	recovery_vault_menu }
}


Function vmBasedOutput([psobject]$localprotecteditem) {
	$vmfound=$false
	for ($i=0; $i -lt $localprotecteditem.count; $i++ ){
		if( $($localprotecteditem[$i].RecoveryAzureVMName).tolower() -eq $vm.tolower() ) 
			{
			$vmfound = $true
			$out1=new-object psobject
			$out1|add-member noteproperty "VM_Name" "$($localprotecteditem[$i].RecoveryAzureVMName)"
			$out1|add-member noteproperty "VM_sku" "$($localprotecteditem[$i].RecoveryAzureVMSize)"
			$out1|add-member noteproperty "Target_ResourceGroup" "$($localprotecteditem[$i].RecoveryResourceGroupId.split('/')[4])"
			
			if($($localprotecteditem[$i].ReplicationHealth) -eq "Normal") 
				{
					if(-not $Grid) { 	$repHealth = "$green Healthy $white"  }
					else { $repHealth = "Healthy" }
				}
			else 
				{
					if (-not $Grid) {	$repHealth = "$red$($localprotecteditem[$i].ReplicationHealth) $white" }
					else{	$repHealth = "$($localprotecteditem[$i].ReplicationHealth)" }
				}
			
			$out1|add-member noteproperty "Replication_health" "$repHealth"
			
			if ($($localprotecteditem[$i].ProtectionState) -eq "Protected" ) 
				{ 
					if(-not $Grid) {	$repProtection = "$green$($localprotecteditem[$i].ProtectionState)$white" }
					else {	$repProtection = "$($localprotecteditem[$i].ProtectionState)" }
				}
			else 
				{
					if (-not $Grid) {	$repProtection = "$red$($localprotecteditem[$i].ProtectionState)$white" }
					else {	$repProtection = "$($localprotecteditem[$i].ProtectionState)" }
				}
			
			$out1|add-member noteproperty "Protection_state" "$repProtection"
			$out1|add-member noteproperty "Target_Vnet" "$($localprotecteditem[$i].SelectedRecoveryAzureNetworkId.split('/')[8])"
			$out1|add-member noteproperty "Target_Subnet" "$($localprotecteditem[$i].NicDetailsList.IpConfigs.RecoverySubnetName)"
			$out1|add-member noteproperty "Target_IP_type" "$($localprotecteditem[$i].NicDetailsList.IpConfigs.RecoveryIPAddressType)"
			$out1|add-member noteproperty "Target_IP_address" "$($localprotecteditem[$i].NicDetailsList.IpConfigs.RecoveryStaticIPAddress)"
			$out1|add-member noteproperty "Storage_Account" "$($localprotecteditem[$i].ProviderSpecificDetails.RecoveryAzureLogStorageAccountId.split('/')[8])"
			$out1|add-member noteproperty "OS_LicenseType" "$($localprotecteditem[$i].ProviderSpecificDetails.LicenseType)"
			
			
			if ($($localprotecteditem[$i].ProviderSpecificDetails.SqlServerLicenseType) -eq "NoLicenseType" )
				{ 
					$sqlahb = "No"
				}
			else 
				{ 
					if (-not $Grid) {	$sqlahb = "${green}Yes${white}" }
					else {	$sqlahb = "Yes" }					
				}
				
			$out1|add-member noteproperty "SQL_AHB" "$sqlahb"
			
			$tags = $localprotecteditem[$i].ProviderSpecificDetails.RecoveryVmTag
			
			$s = ""
			
			foreach ($k in $tags.keys) 
			{
				if ($s) {
					$s += ', [' + $k + ':' + $tags.${k} + ']'
				}
				else {
					$s += '[' + $k + ':' + $tags.${k} + ']'
				}
			}
			
			$out1|add-member noteproperty "Tags" "$s"
				
			$global:out += $out1
			break
			}
		
	}
	
	if(-not $vmfound) {
		Write-Host "Virtual Machine `"$vm`" is not protected under this recovery service vault or subscription.`n" -BackgroundColor Red	
		exit
	}	
	



}


Function regularoutput([psobject]$localprotecteditem) {

		for ($i=0; $i -lt $localprotecteditem.count; $i++ ){
	
		$out1=new-object psobject
		$out1|add-member noteproperty "VM_Name" "$($localprotecteditem[$i].RecoveryAzureVMName)"
		
		if($($localprotecteditem[$i].ReplicationHealth) -eq "Normal") 
			{
				if(-not $Grid) { 	$repHealth = "$green Healthy $white"  }
				else { $repHealth = "Healthy" }
			}
		else 
			{
				if (-not $Grid) {	$repHealth = "$red$($localprotecteditem[$i].ReplicationHealth) $white" }
				else{	$repHealth = "$($localprotecteditem[$i].ReplicationHealth)" }
			}
		
		$out1|add-member noteproperty "Replication_health" "$repHealth"
		
		if ($($localprotecteditem[$i].protectionState) -eq "protected" ) 
			{ 
				if(-not $Grid) {	$repprotection = "$green$($localprotecteditem[$i].protectionState)$white" }
				else {	$repprotection = "$($localprotecteditem[$i].protectionState)" }
			}
		else 
			{
				if (-not $Grid) {	$repprotection = "$red$($localprotecteditem[$i].protectionState)$white" }
				else {	$repprotection = "$($localprotecteditem[$i].protectionState)" }
			}
		
		$out1|add-member noteproperty "protection_state" "$repprotection"
		$out1|add-member noteproperty "Target_Vnet" "$($localprotecteditem[$i].SelectedRecoveryAzureNetworkId.split('/')[8])"
		$out1|add-member noteproperty "Target_Subnet" "$($localprotecteditem[$i].NicDetailsList.IpConfigs.RecoverySubnetName)"

		$out1|add-member noteproperty "VM_sku" "$($localprotecteditem[$i].RecoveryAzureVMSize)"
		$out1|add-member noteproperty "Target_ResourceGroup" "$($localprotecteditem[$i].RecoveryResourceGroupId.split('/')[4])"

		$out1|add-member noteproperty "Target_IP_type" "$($localprotecteditem[$i].NicDetailsList.IpConfigs.RecoveryIPAddressType)"
		$out1|add-member noteproperty "Target_IP_address" "$($localprotecteditem[$i].NicDetailsList.IpConfigs.RecoveryStaticIPAddress)"
		$out1|add-member noteproperty "OS_LicenseType" "$($localprotecteditem[$i].ProviderSpecificDetails.LicenseType)"
		
		
		if ($($localprotecteditem[$i].ProviderSpecificDetails.SqlServerLicenseType) -eq "NoLicenseType" ){ $sqlahb = "No"}
		else { 
				if (-not $Grid) {	$sqlahb = "${green}Yes${white}" }
				else {	$sqlahb = "Yes" }
			
		}
		$out1|add-member noteproperty "SQL_AHB" "$sqlahb"
		$out1|add-member noteproperty "Storage_Account" "$($localprotecteditem[$i].ProviderSpecificDetails.RecoveryAzureLogStorageAccountId.split('/')[8])"
		$tags = $localprotecteditem[$i].ProviderSpecificDetails.RecoveryVmTag
		
		$s = ""
		
		foreach ($k in $tags.keys) 
		{
			if ($s) {
				$s += ', [' + $k + ':' + $tags.${k} + ']'
			}
			else {
				$s += '[' + $k + ':' + $tags.${k} + ']'
			}
		}
		
		$out1|add-member noteproperty "Tags" "$s"
			
		$global:out += $out1
	
	}
	

	
}


Function gridoutput([psobject]$output) {
	Write-Host "`nDetails are presented in another window.`n "
	$output|out-gridview -Title "${global:rsv} : Replicated Items"
}


Function tableoutput([psobject]$output) {
	if((-not $subs) -or (-not $rsv)) {
	Write-Host "`nSubscription :".padright(30) -nonewline ; Write-Host "$global:subscriptionName" -ForegroundColor Green
	Write-Host "Recovery Vault :".padright(30) -nonewline ; Write-Host "$($global:rsv1.name)" -ForegroundColor Green
	}
	
	if($full) { $output|ft * }
	else { $output|ft }
}

Function Moduleinstaller() {
			Write-Host "`"Azure Az`" module is required to run this script. Do you want to install it (Y/N) ? " -Nonewline 
			[char]$response = Read-Host " "
			
			if ($response -match '[Yy]') { 
				Write-Host "Powershell version".padright(20) ":".padleft(5) -nonewline ; Write-Host " $($PSVersionTable.PSVersion)" -ForegroundColor Green
				Write-Host "`nIntializing Az module installer... "
				Write-Host "`nInstalling Az module. Please wait for completion" ; 
				Start-Process PowerShell -Verb RunAs " -ExecutionPolicy RemoteSigned -Command `" Write-Host `"Starting installer`" ;  Install-Module -Name Az -Repository PSGallery -Force`"" -Wait
			}
			else {
				Write-Host "Please install Az module to continue with the script" -BackgroundColor Red
				exit
			}
	
	Modulechecker

}


Function Modulechecker {
	#Check if required Modules are installed.
	$Modulecheck = Get-InstalledModule -Name Az.RecoveryServices -ErrorAction silentlycontinue
	

	if(-not $ModuleCheck) { 
		if($global:ranmodulecheck) {
		Write-Host "Script failed to install Az Module. Please install it manually" -BackgroundColor Red
		Write-Host "Please follow URL https://docs.microsoft.com/en-us/powershell/azure/install-az-ps?view=azps-7.5.0 for instructions" -BackgroundColor Red
		exit
		}
		else {
			$global:ranmodulecheck = $true 
			}
		
		Moduleinstaller 
	}
}


##########
##  Main  ##
##########

$global:ranmodulecheck = $false
$global:out = @()
Modulechecker

# Check if user is logged in to Azure
Write-Host "Checking connectivity to Azure"
if ( ! $(Get-AzContext) )
{
	Write-Host "Failed to connect to Azure. Initiating Azure login..." -BackgroundColor  Red
	Connect-AzAccount
	if (! $?) { exit }
}



if($subs) {
	
	Header; Write-Host "`nSubscription".padright(30) ":".padright(5) -nonewline ; write-host "$subs" -ForegroundColor Green
	set-AzContext -Subscription $subs |out-null
	$global:subscriptionName = $subs

	if ($rsv) { 
		Write-Host "`nRecovery Service Vault".padright(30) ":".padright(5) -nonewline ; write-host "$rsv" -ForegroundColor Green
		$recoveryvault = Get-AzRecoveryServicesVault -Name $rsv 
		
		if(-not $recoveryvault) { 
			Write-Host "`nRecovery Vault $rsv does not exist in $global:subscriptionName subscription. Please check the name and try again" -BackgroundColor Red
			exit 1
		}
	
		Set-AzRecoveryServicesAsrVaultContext -Vault $recoveryvault |out-null
		
		getRSVComponents
	}
	else {	
		recovery_vault_menu 
	}

}
else {
	subscriptionMenu
}

	if($grid) { gridoutput($global:out) }
	else { tableoutput($global:out) }