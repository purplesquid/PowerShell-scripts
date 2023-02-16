#Finds logged in user and gets SID value
#Searches loadbehavior values in the registry for office addins
function Addin{
param([string]$addon, [string]$computerName)
	
	$data = Invoke-Command -ComputerName $computerName -ArgumentList $addon -ScriptBlock{
		param([string]$name)
		
		#Gets active logged in user and removes empty space in front of username
		try{
			$user = (query user | where{$_ -match "Active"}).trim();
		}
		#If no active users, exit
		catch{
			break;
		}
		
		#Get username field (Running remotely removes > in front of username)
		$account = ($user -split "\s+")[0];
		
		#Create ntaccount object and get sid of username
		$objUser = New-Object System.Security.Principal.NTAccount($account); 
		$sid = $objUser.Translate([System.Security.Principal.SecurityIdentifier]); 
		
		$hku = Get-ItemProperty "Microsoft.PowerShell.Core\Registry::HKEY_USERS\$sid\Software\Microsoft\Office\$name\Addins\*"
		$hklmwow = Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\$name\Addins\*"
		
		if($name -eq "Outlook"){
			$hklm = Get-ItemProperty "HKLM:\Software\Microsoft\Office\$name\Addins\*"
		}
		
		$office = @{}
		$office.hkuser = $hku
		$office.hklmwow = $hklmwow
		$office.hklm = $hklm
		$office.user = $user
		$office.error = $error
		$office.sid = $sid
		$office.name = $name
		New-Object psobject -property $office
	}
	
	#If no active users on the machine
	if("$($data.user)" -eq ""){
		Write-Output "No active users logged in or could not connect to the machine"
	}
	else{
		Write-Output "`r`nHKEY_Users\$($data.sid.value)\Software\Microsoft\Office\$($data.name)\Addins"
		Write-Output $data.hkuser | select FriendlyName, PSChildName, LoadBehavior
		Write-Output "`r`n"
	
		Write-Output "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\$($data.name)\Addins`r`n"
		Write-Output $data.hklmwow | select FriendlyName, PSChildName, LoadBehavior
		
		if($addon -eq "Outlook"){
			Write-Output "`r`n`nHKLM:\Software\Microsoft\Office\$($data.name)\Addins`r`n"
			Write-Output $data.hklm| select FriendlyName, PSChildName, LoadBehavior
		}
	}
}
