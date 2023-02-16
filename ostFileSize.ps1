#Gets size of OST files on machine
Function OST{
Param ($device)
	Invoke-Command -ComputerName $device -ScriptBlock{
		#Gets active logged in user and removes empty space in front of username
		try{
			$user = (query user | where{$_ -match "Active"}).trim();
			
			#Get username field (Running remotely removes > in front of username)
			$username = ($user -split "\s+")[0];
			$ost = gci C:\Users\$username\AppData\Local\Microsoft\Outlook\*.ost
			$element = 0

			foreach($i in $ost){
				#Gets ost size in GB
				$size = [double][Math]::Round(($i.Length)/[Math]::Pow(1024,3), 3)
				$lastEdited = $ost.LastWriteTime[$element]
				#Checks if number of ost files equals 1 or more
				if($ost.Length -eq 1){
					$name = $ost.Name
				}
				else{
					$name = $ost[$element]
					$element = $element + 1
				}
				Write-Host "Location:" $name
				Write-Host "Size:" $size "GB"
				Write-Host "Last Write Time:" $lastEdited "`n"
			}
		}
		#No active users logged into the machine
		catch{
			Write-Output "No active users logged in"
		}
		
	}
}
