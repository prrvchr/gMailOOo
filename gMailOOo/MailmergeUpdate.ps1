Param(
    [string] $source,
    [string] $target,
    [switch] $run
)
if($run)
{
    Copy-Item $source $target
}
else
{
    Start-Process -FilePath MailmergeUpdate.ps1 -Verb RunAs -ArgumentList "-source $source -target $target -run"
}
