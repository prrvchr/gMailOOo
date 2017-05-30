Param(
    [string]$source,
    [string]$target
)
Start-Process cmd.exe -ArgumentList "/C move `"$target`" `"$target.bak`" & copy `"$source`" `"$target`"" -Verb RunAs
