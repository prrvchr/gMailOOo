Param(
    [string]$script,
    [string]$source,
    [string]$target
)
Start-Process -FilePath "$script" -Verb RunAs -ArgumentList "$source", "$target"
Start-Sleep -Seconds 60
