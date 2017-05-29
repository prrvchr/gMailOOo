Param(
    [string]$source,
    [string]$target
)
Start-Process -FilePath MailmergeUpdate.cmd -Verb RunAs -ArgumentList "$source", "$target"
