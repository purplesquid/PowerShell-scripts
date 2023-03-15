$word = New-Object -ComObject Word.Application
$path = ""

$files = gci $path\* -Include *.docx, *.doc

$findText = ""
$replaceText = ""

$matchCase = $false
$matchWholeWorld = $true
$matchWildcards = $false
$matchSoundsLike = $false
$matchAllWordForms = $false
$forward = $true
$wrap = 1
$format = $false
$replaceAll = 2


Write-Output "`n$($findText) has been updated with $($replaceText) for the below files:`r`n"

foreach($file in $files){
	#Open each document in folder
	$document = $word.Documents.Open($file.FullName, $null, $false, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true)

	#Find and replace text
	$checkFile = $document.Content.Find.Execute($findText, $matchCase, $matchWholeWorld, $matchWildcards, $matchSoundsLike, $matchAllWordForms, $forward, $wrap, $format, $replaceText, $replaceAll)
	
	#Outputs file name if it finds text and it has updated the file
	if($checkFile -and $document.Saved -eq $false){
		Write-Output $file.FullName
	}

	#Save and close the document
	$document.Close(-1)
}

$word.Quit()
$word = $null
[GC]::Collect()
