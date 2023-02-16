#Compares AD groups for two users
#Retrieves similar and different AD groups between two users

Function CompareGroups{
Param([string]$user1, [string]$user2, [string]$server)

	$reference = Get-ADPrincipalGroupMembership -server server -Identity $user1
	$difference = Get-ADPrincipalGroupMembership -server server -Identity $user2
	$differentGroups = compare-object -referenceobject $reference -differenceobject $difference -Property samaccountname | sort samaccountname
	$equalGroups = compare-object -referenceobject $reference -differenceobject $difference -IncludeEqual -ExcludeDifferent -Property samaccountname | sort SamAccountName
	$userarraylist1 = New-Object -TypeName "System.Collections.ArrayList"
	$userarraylist2 = New-Object -TypeName "System.Collections.ArrayList"

	foreach($group in $differentGroups){
		if($group.SideIndicator -eq "<="){
			#cast void to prevent index being returned
			[void]$userarraylist1.Add($group.samaccountname)
		}
		else{
			[void]$userarraylist2.Add($group.samaccountname)
		}
	}
	
	Write-Output "`n"
	Write-Output $user1
	Write-Output $userarraylist1
	Write-Output "`n"
	Write-Output $user2
	Write-Output $userarraylist2
	Write-Output "`n"
	
	Write-Output "Groups that both users share"
	Write-Output $equalGroups.samaccountname
}
