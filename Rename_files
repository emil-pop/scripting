Get-childItem -Filter "SG1 [5*" -Recurse | Rename-Item -NewName {$_.name -replace '(\[5-\d\d)', '$1]'}
