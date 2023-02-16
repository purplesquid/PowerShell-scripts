Function memberOf{
Param ($group, $domain)
	
	$ldap = [adsi] "LDAP://$domain"
	$searcher = new-object System.DirectoryServices.DirectorySearcher($ldap)
	$searcher.filter = "name=$group"
	$members = $searcher.FindOne().Properties.memberof | sort
	
	foreach($member in $members){
		$name = ($member -split("="))[1] 
		$name = $name.Replace(',OU','') 
		$name = $name.Replace('\','')
		$email = ([adsisearcher]"name=$($name)").FindOne().Properties.mail
		
		Write-Output "$($name);$($email)";		
	}
}
