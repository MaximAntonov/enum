# Description: this script finds keywords in MS Word files (.doc and .docs)

# Source:https://stackoverflow.com/questions/58322328/powershell-find-ms-word-files-through-keywords

# Caveat: if the script discovers a password-protected file, it stops and waits for the password input, and then opens up other ms word documents for a fraction of a second

# Usage: .\wordpass.ps1

add-type -AssemblyName "Microsoft.Office.Interop.Word"

#Folder to connect to
$SourceFolder = "c:\users\$env:UserName\"
cd $SourceFolder

#Keywords to search for
$keyword1 = "пароль"
$Forward = $true
$MatchWholeWord = $true

$Word = New-Object -ComObject Word.Application
$docs = Get-ChildItem -Path $SourceFolder -Include @("*.doc", "*.docx") -Recurse

foreach ($doc in $docs)
{
    $condition1 = $Word.Documents.Open($doc.FullName).Content.Find.Execute($keyword1,$MatchWholeWord)

    switch($condition1)
    {
        $true
        {
            #$word.Application.ActiveDocument.Close()
            Write-Host -f Cyan "$doc contains the keyword: '$keyword1'"
            #Move-Item -Path $doc.FullName -Destination $destination
            $word.Application.ActiveDocument.Close()
        }

        $false
        {
            $word.Application.ActiveDocument.Close()
            #Write-Host -f Red "$doc does not contain the keyword: $keyword1"
        }
    }

    Write-Host "Filename '$($doc.Fullname)"
    Write-Host "`r"

}

Stop-Process -Name "WINWORD"
