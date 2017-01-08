function calcPdfPTableHeight
{
    param
    (
        [Parameter(Mandatory = $true)][iTextSharp.text.pdf.PDFPTable]$Tab
    )

    $ms = New-Object -TypeName System.IO.MemoryStream
    $virtPdf = New-Object -TypeName iTextSharp.text.Document
    $virtPdf.SetPageSize([iTextSharp.text.PageSize]::A4) | Out-Null
    $wr = [iTextSharp.text.pdf.PdfWriter]::GetInstance($virtPdf, $ms)

    $virtPdf.Open()
    $tab.WriteSelectedRows(0, $tab.Rows.Count, 0, 0, $wr.DirectContent) | Out-Null
    $virtPdf.Close()

    return $tab.TotalHeight
}

function genPageHeader
{
    param
    (
        [Parameter(Mandatory = $true)][iTextSharp.text.Font]$Font,
        [Parameter(Mandatory = $true)][iTextSharp.text.BaseColor]$BorderColor,
        [Parameter(Mandatory = $true)][System.String]$LogoPath,
        [Parameter(Mandatory = $true)][System.String]$ProjName,
        [Parameter(Mandatory = $true)][System.String]$ProjVer
    )

    $logo = [iTextSharp.text.Image]::GetInstance($LogoPath)
    $logo.ScalePercent(25)

    $tab = New-Object -TypeName iTextSharp.text.pdf.PDFPTable(2)
    $tab.WidthPercentage = 100
    $tab.SpacingBefore = 0.0
    $cell = New-Object -TypeName iTextSharp.text.pdf.PdfPCell
    $cell.AddElement($logo) | Out-Null
    $cell.Border = [iTextSharp.text.Rectangle]::BOTTOM_BORDER
    $cell.BorderWidth = 2.0
    $cell.BorderColor = $BorderColor
    $cell.HorizontalAlignment = [iTextSharp.text.Element]::ALIGN_LEFT
    $tab.AddCell($cell) | Out-Null

    $nestTab = New-Object -TypeName iTextSharp.text.pdf.PDFPTable(1)
    $nestTab.WidthPercentage = 100
    $cell = New-Object -TypeName iTextSharp.text.pdf.PdfPCell
    $cell.PaddingBottom = -2.0

    $p = New-Object -TypeName iTextSharp.text.Paragraph
    $p.Font = $Font
    $p.SpacingBefore = 0.0
    $p.SpacingAfter = 0.0
    $p.Add($ProjName) | Out-Null
    $p.Alignment = [iTextSharp.text.Element]::ALIGN_RIGHT

    $cell.AddElement($p) | Out-Null
    $cell.Border = [iTextSharp.text.Rectangle]::NO_BORDER
    $cell.HorizontalAlignment = [iTextSharp.text.Element]::ALIGN_RIGHT
    $nestTab.AddCell($cell) | Out-Null

    $cell = New-Object -TypeName iTextSharp.text.pdf.PdfPCell

    $p = New-Object -TypeName iTextSharp.text.Paragraph
    $p.Font = $Font
    $p.SpacingBefore = 0.0
    $p.SpacingAfter = 0.0
    $p.Add($ProjVer) | Out-Null
    $p.Alignment = [iTextSharp.text.Element]::ALIGN_RIGHT

    $cell.AddElement($p) | Out-Null

    $cell.Border = [iTextSharp.text.Rectangle]::NO_BORDER
    $cell.HorizontalAlignment = [iTextSharp.text.Element]::ALIGN_RIGHT
    $nestTab.AddCell($cell) | Out-Null

    $cell = New-Object -TypeName iTextSharp.text.pdf.PdfPCell
    $cell.Border = [iTextSharp.text.Rectangle]::BOTTOM_BORDER
    $cell.BorderWidth = 2.0
    $cell.BorderColor = $BorderColor
    $cell.AddElement($nestTab) | Out-Null
    $tab.AddCell($cell) | Out-Null

    $cell = New-Object -TypeName iTextSharp.text.pdf.PdfPCell

    $p = New-Object -TypeName iTextSharp.text.Paragraph
    $p.Font = $Font
    $p.SpacingBefore = 0.0
    $p.SpacingAfter = 0.0
    $p.Add("Описание задач на изменение") | Out-Null
    $p.Alignment = [iTextSharp.text.Element]::ALIGN_LEFT

    $cell.AddElement($p) | Out-Null
    $cell.Border = [iTextSharp.text.Rectangle]::NO_BORDER
    $cell.HorizontalAlignment = [iTextSharp.text.Element]::ALIGN_LEFT
    $tab.AddCell($cell) | Out-Null

    $cell = New-Object -TypeName iTextSharp.text.pdf.PdfPCell

    $p = New-Object -TypeName iTextSharp.text.Paragraph
    $p.Font = $Font
    $p.SpacingBefore = 0.0
    $p.SpacingAfter = 0.0
    $p.Add("Дата отчёта: $(Get-Date -Format "dd.MM.yyyy")") | Out-Null
    $p.Alignment = [iTextSharp.text.Element]::ALIGN_RIGHT

    $cell.AddElement($p) | Out-Null
    $cell.Border = [iTextSharp.text.Rectangle]::NO_BORDER
    $cell.HorizontalAlignment = [iTextSharp.text.Element]::ALIGN_RIGHT
    $tab.AddCell($cell) | Out-Null

    return $tab
}

function genTabHeader
{
    param
    (
        [Parameter(Mandatory = $true)][iTextSharp.text.Font]$BdFont,
        [Parameter(Mandatory = $true)][iTextSharp.text.BaseColor]$BorderColor
    )

    $tab = New-Object -TypeName iTextSharp.text.pdf.PDFPTable(2)
    $tab.TotalWidth = 555.0
    $tab.WidthPercentage = 100
    [float[]]$Widths = @(120.0, 435.0)
    $tab.SetWidths($Widths)
    $tab.SpacingBefore = 0.0
    $tab.SpacingAfter = 0.0
    $cell = New-Object -TypeName iTextSharp.text.pdf.PdfPCell

    $p = New-Object -TypeName iTextSharp.text.Paragraph
    $p.Font = $BdFont
    $p.SpacingBefore = 2.0
    $p.SpacingAfter = 6.0
    $p.Add("Задача") | Out-Null

    $cell.AddElement($p) | Out-Null
    $cell.Border = [iTextSharp.text.Rectangle]::BOTTOM_BORDER -bor [iTextSharp.text.Rectangle]::TOP_BORDER
    $cell.BorderWidth = 2.0
    $cell.BorderColor = $BorderColor
    $cell.PaddingLeft = 10
    $tab.AddCell($cell) | Out-Null

    $cell = New-Object -TypeName iTextSharp.text.pdf.PdfPCell

    $p = New-Object -TypeName iTextSharp.text.Paragraph
    $p.Font = $BdFont
    $p.SpacingBefore = 2.0
    $p.SpacingAfter = 6.0
    $p.Add("Описание") | Out-Null

    $cell.AddElement($p) | Out-Null
    $cell.Border = [iTextSharp.text.Rectangle]::BOTTOM_BORDER -bor [iTextSharp.text.Rectangle]::TOP_BORDER
    $cell.BorderWidth = 2.0
    $cell.BorderColor = $BorderColor
    $cell.PaddingLeft = 10
    $tab.AddCell($cell) | Out-Null

    return $tab
}

function genPdfReport
{
    param
    (
        [Parameter(Mandatory = $true)][System.String]$Path,
        [Parameter(Mandatory = $true)][System.String]$ProjKey,
        [Parameter(Mandatory = $true)][System.String]$ProjVer,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSObject[]]$Issues
    )

    switch($ProjKey)
    {
        "RETAIL" { $projName = "PSB-Retail" }
        "NA" { $projName = "Новая Афина" }
        "DIAUP" { $projName = "Diasoft" }
        "FACT" { $projName = "Факторинг" }
        default {$projName = "Unknown"}
    }

    #$scriptRoot = "C:\Users\NEX\Documents\jk10\retail_scripts"
    #$scriptRoot = "D:\Retail_scripts"
    $tabMaxHeight = 680
    $headerHeight = 90.0
    $bodyHeight = 682.0
    $footerHeight = 30.0
    $pageNum = 1
    #$logoPath = "$scriptRoot\PSB_logo_ru_hor2112\PSB_logo_ru_hor.gif"
    #Add-Type -Path "$scriptRoot\itextsharp-all-5.5.10\itextsharp-dll-core\itextsharp.dll"
    $logoPath = "$PSScriptRoot\PSB_logo_ru_hor2112\PSB_logo_ru_hor.gif"
    Add-Type -Path "$PSScriptRoot\itextsharp-all-5.5.10\itextsharp-dll-core\itextsharp.dll"

    $pdf = New-Object -TypeName iTextSharp.text.Document
    $pdf.SetPageSize([iTextSharp.text.PageSize]::A4) | Out-Null
    $pdf.SetMargins(20.0, 20.0, 20.0, 20.0) | Out-Null
    $pdf.AddAuthor("someone") | Out-Null
    #[void][iTextSharp.text.pdf.PdfWriter]::GetInstance($pdf, [System.IO.File]::Create("$scriptRoot\retail_report.pdf"))
    [void][iTextSharp.text.pdf.PdfWriter]::GetInstance($pdf, [System.IO.File]::Create($Path))

    $PSBBlue = New-Object iTextSharp.text.BaseColor(5, 64, 142)
    $PSBOrange = New-Object iTextSharp.text.BaseColor(233, 93, 15)
    #$arialBaseFont = [iTextSharp.text.pdf.BaseFont]::CreateFont("$scriptRoot\fonts\arial_fonts\ARIAL.TTF", [iTextSharp.text.pdf.BaseFont]::IDENTITY_H, [iTextSharp.text.pdf.BaseFont]::EMBEDDED)
    $arialBaseFont = [iTextSharp.text.pdf.BaseFont]::CreateFont("$PSScriptRoot\fonts\arial_fonts\ARIAL.TTF", [iTextSharp.text.pdf.BaseFont]::IDENTITY_H, [iTextSharp.text.pdf.BaseFont]::EMBEDDED)
    $arialFont = New-Object iTextSharp.text.Font($arialBaseFont, 12, [iTextSharp.text.Font]::NORMAL, $PSBBlue)
    #$arialbdBaseFont = [iTextSharp.text.pdf.BaseFont]::CreateFont("$scriptRoot\fonts\arial_fonts\ARIALBD.TTF", [iTextSharp.text.pdf.BaseFont]::IDENTITY_H, [iTextSharp.text.pdf.BaseFont]::EMBEDDED)
    $arialbdBaseFont = [iTextSharp.text.pdf.BaseFont]::CreateFont("$PSScriptRoot\fonts\arial_fonts\ARIALBD.TTF", [iTextSharp.text.pdf.BaseFont]::IDENTITY_H, [iTextSharp.text.pdf.BaseFont]::EMBEDDED)
    $arialbdFont = New-Object iTextSharp.text.Font($arialbdBaseFont, 14, [iTextSharp.text.Font]::NORMAL, $PSBBlue)

    $pdf.Open()

    $pageLayout = New-Object -TypeName iTextSharp.text.pdf.PDFPTable(1)
    $pageLayout.TotalWidth = 555.0
    $pageLayout.WidthPercentage = 100
    $pageLayout.SpacingBefore = 0.0
    $pageLayout.SpacingAfter = 0.0

    $header = New-Object -TypeName iTextSharp.text.pdf.PdfPCell
    $header.Border = [iTextSharp.text.Rectangle]::NO_BORDER
    $header.FixedHeight = $headerHeight
    $tab = genPageHeader -Font $arialFont -BorderColor $PSBBlue -LogoPath $logoPath -ProjName $projName -ProjVer $ProjVer
    $header.AddElement($tab) | Out-Null
    $pageLayout.AddCell($header) | Out-Null

    $body = New-Object -TypeName iTextSharp.text.pdf.PdfPCell
    $body.Border = [iTextSharp.text.Rectangle]::NO_BORDER
    $body.FixedHeight = $bodyHeight

    $tab = genTabHeader -BdFont $arialbdFont -BorderColor $PSBOrange
    $auxTab = genTabHeader -BdFont $arialbdFont -BorderColor $PSBOrange

    $prevTaskCell = $null
    $prevDescCell = $null

    foreach($task in $Issues)
    {
        $taskCell = New-Object -TypeName iTextSharp.text.pdf.PdfPCell

        $p = New-Object -TypeName iTextSharp.text.Paragraph
        $p.Font = $arialFont
        $p.SpacingBefore = 2.0
        $p.SpacingAfter = 6.0
        $p.Add($task.Key) | Out-Null

        $taskCell.AddElement($p) | Out-Null
        $taskCell.Border = [iTextSharp.text.Rectangle]::BOTTOM_BORDER
        $taskCell.BorderWidth = 1.0
        $taskCell.BorderColor = $PSBOrange
        $taskCell.PaddingLeft = 10
        $auxTab.AddCell($taskCell) | Out-Null

        $descCell = New-Object -TypeName iTextSharp.text.pdf.PdfPCell

        $p = New-Object -TypeName iTextSharp.text.Paragraph
        $p.Font = $arialFont
        $p.SpacingBefore = 2.0
        $p.SpacingAfter = 6.0
        $p.Add($task.Sum) | Out-Null

        $descCell.AddElement($p) | Out-Null
        $descCell.Border = [iTextSharp.text.Rectangle]::BOTTOM_BORDER
        $descCell.BorderWidth = 1.0
        $descCell.BorderColor = $PSBOrange
        $descCell.PaddingLeft = 10
        $auxTab.AddCell($descCell) | Out-Null

        if((calcPdfPTableHeight -Tab $auxTab) -gt $tabMaxHeight)
        {
            if(! ($prevTaskCell -and $prevDescCell))
            {
                break
            }

            $prevTaskCell.BorderWidth = 2.0
            $prevDescCell.BorderWidth = 2.0

            $tab.AddCell($prevTaskCell) | Out-Null
            $tab.AddCell($prevDescCell) | Out-Null

            $body.AddElement($tab) | Out-Null
            $pageLayout.AddCell($body) | Out-Null

            $footer = New-Object -TypeName iTextSharp.text.pdf.PdfPCell
            $footer.Border = [iTextSharp.text.Rectangle]::NO_BORDER
            $footer.FixedHeight = $footerHeight

            $p = New-Object -TypeName iTextSharp.text.Paragraph
            $p.Font = $arialFont
            $p.SpacingBefore = 0.0
            $p.SpacingAfter = 0.0
            $p.Add($pageNum) | Out-Null
            $p.Alignment = [iTextSharp.text.Element]::ALIGN_CENTER
            $pageNum++

            $footer.AddElement($p) | Out-Null
            $pageLayout.AddCell($footer) | Out-Null

            $pdf.Add($pageLayout) | Out-Null

            $pdf.NewPage() | Out-Null

            $pageLayout = New-Object -TypeName iTextSharp.text.pdf.PDFPTable(1)
            $pageLayout.TotalWidth = 555.0
            $pageLayout.WidthPercentage = 100
            $pageLayout.SpacingBefore = 0.0
            $pageLayout.SpacingAfter = 0.0

            $header = New-Object -TypeName iTextSharp.text.pdf.PdfPCell
            $header.Border = [iTextSharp.text.Rectangle]::NO_BORDER
            $header.FixedHeight = $headerHeight
            $tab = genPageHeader -Font $arialFont -BorderColor $PSBBlue -LogoPath $logoPath -ProjName $projName -ProjVer $ProjVer
            $header.AddElement($tab) | Out-Null
            $pageLayout.AddCell($header) | Out-Null

            $body = New-Object -TypeName iTextSharp.text.pdf.PdfPCell
            $body.Border = [iTextSharp.text.Rectangle]::NO_BORDER
            $body.FixedHeight = $bodyHeight

            $tab = genTabHeader -BdFont $arialbdFont -BorderColor $PSBOrange
            $auxTab = genTabHeader -BdFont $arialbdFont -BorderColor $PSBOrange
        
            $auxTab.AddCell($taskCell) | Out-Null
            $auxTab.AddCell($descCell) | Out-Null
        }
        else
        {
            if($prevTaskCell -and $prevDescCell)
            {
                $tab.AddCell($prevTaskCell) | Out-Null
                $tab.AddCell($prevDescCell) | Out-Null
            }
        }

        $prevTaskCell = $taskCell
        $prevDescCell = $descCell
    }

    $prevTaskCell.BorderWidth = 2.0
    $prevDescCell.BorderWidth = 2.0

    $tab.AddCell($prevTaskCell) | Out-Null
    $tab.AddCell($prevDescCell) | Out-Null

    $body.AddElement($tab) | Out-Null
    $pageLayout.AddCell($body) | Out-Null

    $footer = New-Object -TypeName iTextSharp.text.pdf.PdfPCell
    $footer.Border = [iTextSharp.text.Rectangle]::NO_BORDER
    $footer.FixedHeight = $footerHeight

    $p = New-Object -TypeName iTextSharp.text.Paragraph
    $p.Font = $arialFont
    $p.SpacingBefore = 0.0
    $p.SpacingAfter = 0.0
    $p.Add($pageNum) | Out-Null
    $p.Alignment = [iTextSharp.text.Element]::ALIGN_CENTER

    $footer.AddElement($p) | Out-Null
    $pageLayout.AddCell($footer) | Out-Null

    $pdf.Add($pageLayout) | Out-Null

    $pdf.Close()
}

<#
$testArr = @()

for($j = 0; $j -lt 50; $j++)
{
    $obj = New-Object -TypeName psobject -Property @{

        Key = "RETAIL-$(10000 + $j)"
        Sum = "Разработка по заявке RFC-$(1000 + $j)"
    }

    $testArr += $obj
}

genPdfReport -ProjKey "RETAIL" -ProjVer "RC_18.01.2017" -Issues $testArr
#>