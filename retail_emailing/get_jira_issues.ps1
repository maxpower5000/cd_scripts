param
(
    [Parameter(Mandatory = $true)][System.String]$ProjKey,
    [Parameter(Mandatory = $true)][System.String]$ProjVer,
    [Parameter(Mandatory = $true)][System.String]$UserName,
    [Parameter(Mandatory = $true)][System.String]$Password,
    [Parameter(Mandatory = $true)][System.String]$WorkSpace
)

try
{
    $issues = Import-CliXml -Path "$WorkSpace\retail_issues.xml" -ErrorAction Stop
    Export-CliXml -InputObject $issues -Path "$WorkSpace\tmp\jira_issues.xml" -ErrorAction Stop
}
catch
{
    $_ | Out-File -FilePath "$WorkSpace\tmp\ps_errors.txt" -Append
}