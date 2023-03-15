#DCOM and WMI inbound rules needs to be enabled in firewall
#DCOM port 135
function portTest {
	Param(
		[string] $server,
		$port=135, 
		$timeout=3000
	)
	
	$tcpClient = new-Object system.Net.Sockets.TcpClient
	$connection = $tcpClient.BeginConnect($server,$port,$null,$null)
	$wait = $connection.AsyncWaitHandle.WaitOne($timeout,$false)

	if(!$wait){
		Write-Host "$server timed out"
		$tcpClient.close()
		$tcpClient.dispose()
		return $false;
	}
	else{
		Write-Host "Connected to: $server"
		$tcpClient.close()
		$tcpClient.dispose()
	}
}

#List Servers 
$serverNames = ""
$username = ""
$password = ConvertTo-SecureString -String "" -AsPlainText -Force
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password

foreach($server in $serverNames){
	$serverCheck = portTest $server
	
	if($serverCheck -ne "False"){
		$bootUpTime = Get-WmiObject Win32_OperatingSystem -computername $server -credential $credential
		$convertBootDate = $bootUpTime.ConvertToDateTime($bootUpTime.lastbootuptime)
		Write-Host "Last reboot: " $convertBootDate "`n"
		Invoke-command -computername $server -credential $credential -scriptblock{Restart-Computer -Force}
	}
	else{
		Write-Host "Reboot failed for: $server"
	}
}
