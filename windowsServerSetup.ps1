#Testing WinRM port
function portTest {

	Param(
		[string] $server,
		$port=5985, 
		$timeout=3000
	)
	
	$tcpClient = new-Object system.Net.Sockets.TcpClient
	$connection = $tcpClient.BeginConnect($server,$port,$null,$null)
	$wait = $connection.AsyncWaitHandle.WaitOne($timeout,$false)

	if(!$wait){
		Write-Host "WinRM is not allowed on: $($server)"
		$tcpClient.close()
		$tcpClient.dispose()
		return $false;
	}
	else{
		Write-Verbose "Connected to: $server"
		$tcpClient.close()
		$tcpClient.dispose()
	}
}

function qemuagent{
	$username=""
	$password = ConvertTo-SecureString "" -AsPlainText -Force
	$cred = New-Object System.Management.Automation.PSCredential ($username, $password)
	
	$servers = ""
	
	foreach($server in $servers){
		$serverCheck = portTest $server
		
		if($serverCheck -ne "False"){
			Invoke-Command -ComputerName $server -Credential $cred -ScriptBlock{
				$version = wmic os get Caption /value
				$driveLetter = (Get-WmiObject win32_volume | where-object{$_.Label -match "virtio"}).DriveLetter
				
				if($version -match "Windows Server 2008"){
					$serialPath = "vioserial\2k8\amd64"
					$balloonPath = "Balloon\2k8\amd64"
				}
				elseif($version -match "Windows Server 2008 R2"){
					$serialPath = "vioserial\2k8R2\amd64"
					$balloonPath = "Balloon\2k8R2\amd64"
				}
				elseif($version -match "Windows Server 2012"){
					$serialPath = "vioserial\2k12\amd64"
					$balloonPath = "Balloon\2k12\amd64"
				}
				elseif($version -match "Windows Server 2012 R2"){
					$serialPath = "vioserial\2k12R2\amd64"
					$balloonPath = "Balloon\2k12R2\amd64"
				}
				elseif($version -match "Windows Server 2016"){
					$serialPath = "vioserial\2k16\amd64"
					$balloonPath = "Balloon\2k16\amd64"
				}
				elseif($version -match "Windows Server 2019"){
					$serialPath = "vioserial\2k19\amd64"
					$balloonPath = "Balloon\2k19\amd64"
				}
				
				$virtioFullPath = "$driveLetter\$serialPath\vioser.inf"
				
				#Install VirtIO serial driver
				if(Test-Path $virtioFullPath){
					pnputil.exe -a $virtioFullPath /install | Out-Null
								
					if($LASTEXITCODE -eq 1){
						Write-Output "VirtIO serial driver failed to install"
					}
					else{
						Write-Output "VirtIO serial driver installed"
					}
				}
				
				$balloonFullPath = "$driveLetter\$balloonPath\balloon.inf"
				
				#Install VirtIO Balloon driver
				if(Test-Path $balloonFullPath){
					pnputil.exe -a $balloonFullPath /install | Out-Null
								
					if($LASTEXITCODE -eq 1){
						Write-Output "VirtIO Balloon driver failed to install"
					}
					else{
						Write-Output "VirtIO Balloon driver installed"
					}
				}
				
				#QEMU Agent Install
				if([Environment]::Is64BitOperatingSystem){
					Start-Process msiexec.exe -Wait -ArgumentList "/I $driveLetter\guest-agent\qemu-ga-x86_64.msi /quiet"
					
					if($LASTEXITCODE -eq 1){
						Write-Output "VirtIO Balloon driver failed to install"
					}
					else{
						Write-Output "QEMU Guest Agent installed"
					}
				}
				else{
					Start-Process msiexec.exe -Wait -ArgumentList "/I $driveLetter\guest-agent\qemu-ga-i386.msi /quiet"
					
					if($LASTEXITCODE -eq 1){
						Write-Output "VirtIO Balloon driver failed to install"
					}
					else{
						Write-Output "QEMU Guest Agent installed"
					}
				}
				
				#Enable ICMP ping
				Write-Output "`r`nEnabling ICMP`r`n"
				Enable-NetFirewallRule -DisplayName "File and Printer Sharing (Echo Request - ICMPv4-In)"
				Enable-NetFirewallRule -DisplayName "File and Printer Sharing (Echo Request - ICMPv6-In)"
				
				#Setting timezone
				$timeZone = "Eastern Standard Time"
				Write-Output "Setting timezone to $($timeZone)`r`n"
				Set-TimeZone -Name $timeZone -PassThru | Out-Null
				
				#Enabling RDP
				Write-Output "Enabling RDP`r`n"
				Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -name "fDenyTSConnections" -value 0
				Enable-NetFirewallRule -DisplayName "Remote Desktop - Shadow (TCP-In)"
				Enable-NetFirewallRule -DisplayName "Remote Desktop - User Mode (TCP-In)"
				Enable-NetFirewallRule -DisplayName "Remote Desktop - User Mode (UDP-In)"
				
				#Joining domain
				$fullDomainName = ""
				$domain = $fullDomainName.split(".")[0]
				$endDomain = $fullDomainName.split(".")[1]
				$domainUser = ""
				$domainPassword = ""
				$OU = "OU=Windows,OU=Servers,DC=$($domain),DC=$($endDomain)"
				
				$joinDomain = (Get-WMIObject -NameSpace "Root\Cimv2" -Class "Win32_ComputerSystem").JoinDomainOrWorkgroup($fullDomainName, $domainPassword, "$($domain)\$($domainUser)", $OU, 3)
				
				if($joinDomain.ReturnValue -eq "2691"){
					Write-Output "$($env:computername) is already joined to $($fullDomainName)"
				}
				else{
					Write-Output "Joining $($fullDomainName)"
					$rebootPrompt = Read-Host "Do you want to reboot now for the domain join to take effect: [y/n]"
          
					if($rebootPrompt -eq "y" -or $rebootPrompt -eq "Y"){
						Restart-Computer -Force
					}
					else{
						"Server will need to be rebooted before $($env:computername) is joined to $($fullDomainName)"
					}
				}
			}
		}
	}
}

qemuagent
