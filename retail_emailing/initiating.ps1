param
(
    [Parameter(Mandatory = $true)][System.String]$WorkSpace
)

try
{
    if(! (Test-Path -Path $WorkSpace))
    {
        throw "Workspace does not exist:`r`n$WorkSpace"
    }
    
    if(Test-Path -Path "$WorkSpace\tmp")
    {
        Get-ChildItem -Path "$WorkSpace\tmp" -ErrorAction Stop | Remove-Item -ErrorAction Stop
    }
    else
    {
        New-Item -Path "$WorkSpace\tmp" -Type directory -ErrorAction Stop
    }
    throw "Custom error"
}
catch
{
    "ERROR: Ошибка при получении списка задач из Jira." | Out-File -FilePath "C:\Program Files (x86)\Jenkins\workspace\retail_emailing\ps_errors.txt" -Append -Encoding utf8
    $_.Exception.Message | Out-File -FilePath "C:\Program Files (x86)\Jenkins\workspace\retail_emailing\ps_errors.txt" -Append -Encoding utf8
    #[Environment]::SetEnvironmentVariable("ps_cust_output", "ERROR: Ошибка при получении списка задач из Jira.\r\n$($_.Exception.Message)", "Machine")
    #Write-Host -Object "ERROR: Ошибка при получении списка задач из Jira." -ForegroundColor Red
    #Write-Host -Object $_.Exception.Message -ForegroundColor Red
    exit 1
}