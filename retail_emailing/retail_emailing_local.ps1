param
(
    [Parameter(Mandatory = $true)][System.String]$ProjKey,
    [Parameter(Mandatory = $true)][System.String]$ProjVer,
    [Parameter(Mandatory = $true)][System.String]$UserName,
    [Parameter(Mandatory = $true)][System.String]$Password
)

."$PSScriptRoot\global_vars.ps1"
."$PSScriptRoot\gen_pdf_report.ps1"

add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;

    public class TrustAllCertsPolicy : ICertificatePolicy
    {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem)
        {
            return true;
        }
    }
"@

[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

$secStrPass = ConvertTo-SecureString -String $Password -AsPlainText -Force
$BSTR = [System.Runtime.InteropServices.marshal]::SecureStringToBSTR($secStrPass)
$JiraPass = [System.Runtime.InteropServices.marshal]::PtrToStringAuto($BSTR)
$JiraLogin = $UserName
$bytes = [System.Text.Encoding]::UTF8.GetBytes("$JiraLogin`:$JiraPass")
$EncodedCredentials = [System.Convert]::ToBase64String($bytes)

$Body =
@{
    jql = "project = $ProjKey AND fixversion = $ProjVer"
    fields = @("issuetype", "assignee", "summary", "issuelinks")
}

$JsonBody = $Body | ConvertTo-Json

try
{
    $issues = Import-CliXml -Path "$PSScriptRoot\retail_issues.xml" -ErrorAction Stop
    $issues = $issues.issues
}
catch
{
    Write-Host -Object "ERROR: Couldn't get the issues list from Jira." -ForegroundColor Red
    Write-Host -Object $_.Exception.Message -ForegroundColor Red
    exit 1
}

$pdfData = @()
$emailData = ""

foreach($task in $issues)
{
    $emailData += "<li><a href=""$jiraUrl/browse/$($task.key)"">$($task.key)</a>: $($task.fields.summary)</li>"
    $emailDataLinkedIssues = ""

    $taskObj = New-Object -TypeName psobject -Property @{
        Key = $task.key
        Sum = $task.fields.summary
    }

    $pdfDataLinkedIssues = @()

    # 10107 - Разработка по заявкам RFC
    if($task.fields.issuetype.id -eq "10107")
    {
        foreach($linkedIssue in $task.fields.issuelinks)
        {
            $inwardIssueKey = $null
            $inwardIssueSum = $null

            $inwardIssueKey = $linkedIssue.inwardIssue.key
            $inwardIssueSum = $linkedIssue.inwardIssue.fields.summary

            if($inwardIssueKey -and $inwardIssueSum)
            {
                $emailDataLinkedIssues += "<li><a href=""$jiraUrl/browse/$inwardIssueKey"">$inwardIssueKey</a>: $inwardIssueSum</li>"
                $pdfDataLinkedIssues += $inwardIssueSum
            }
        }
    }

    if($emailDataLinkedIssues.length -gt 0)
    {
        $emailData += "<ul>$emailDataLinkedIssues</ul>"
    }

    if($pdfDataLinkedIssues.Count -gt 0)
    {
        $taskObj.Sum += " - " + ($pdfDataLinkedIssues -join ", ")
    }

    $pdfData += $taskObj
}

if($emailData.length -gt 0)
{
    $emailData = "<ol>$emailData</ol>"
}

if(Test-Path -Path "$PSScriptRoot\$pdfReport")
{
    try
    {
        Remove-Item -Path "$PSScriptRoot\$pdfReport" -ErrorAction Stop
    }
    catch
    {
        Write-Host -Object "ERROR: Couldn't remove the old pdf report." -ForegroundColor Red
        Write-Host -Object $_.Exception.Message -ForegroundColor Red
    }
}

$Subject += $ProjVer
$attachPdfReport = $true
$Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $secStrPass

try
{
    genPdfReport -Path "$PSScriptRoot\$pdfReport" -ProjKey $ProjKey -ProjVer $ProjVer -Issues $pdfData
}
catch
{
    Write-Host -Object "ERROR: Couldn't generate a pdf report." -ForegroundColor Red
    Write-Host -Object $_.Exception.Message -ForegroundColor Red
    $attachPdfReport = $false
}

try
{
    if($attachPdfReport)
    {
        Send-MailMessage -From $From -To $To -Subject $Subject -Body $emailData -BodyAsHtml -SmtpServer $SmtpServer -Credential $Cred -Attachments "$PSScriptRoot\$pdfReport" `
            -Encoding $([System.Text.Encoding]::UTF8) -UseSsl -ErrorAction Stop
    }
    else
    {
        Send-MailMessage -From $From -To $To -Subject $Subject -Body $emailData -BodyAsHtml -SmtpServer $SmtpServer -Credential $Cred `
            -Encoding $([System.Text.Encoding]::UTF8) -UseSsl -ErrorAction Stop
    }
}
catch
{
    Write-Host -Object "ERROR: Couldn't send email." -ForegroundColor Red
    Write-Host -Object $_.Exception.Message -ForegroundColor Red
    exit 1
}