$directory = '\\Network\Path\To\Documents'
$thing2find = 'did jj tie buckle'
$files = Get-ChildItem $directory -Include *.doc,*.docx -Recurse
$application = New-Object -ComObject word.application
$application.visible = $false
$results = @()

Function getStringMatch{
    Foreach ($file in $files){
        $document = $application.documents.open($file.FullName,$false,$true)
        $contents = $document.content
        If($contents.Text -match "$($thing2find)"){
            $fileName = $file.FullName
            $results += "$fileName `n"
        }
        $document.close()
    }
    If($results){
        Write-Host "Found \"$thing2find\" in the following documents: "
        Write-Host $results
    }
    #$document.close()
    $application.quit()
}
getStringMatch
