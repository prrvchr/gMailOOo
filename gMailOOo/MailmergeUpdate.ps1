$source = $args[0]
$target = $args[1]
Start-Process -FilePath PowerShell.exe -Verb RunAs -ArgumentList @('-Command &{Move-Item $target $target.bak -Force; Copy-Item $source $target -Force}')