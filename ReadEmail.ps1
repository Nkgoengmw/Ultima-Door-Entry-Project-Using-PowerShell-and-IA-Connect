function Get-UsersFromEmail {

    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace('MAPI')
    $inbox = $namespace.GetDefaultFolder(6)

    $email = $inbox.Items |
        Where-Object { $_.Subject -like "*Employee ID*" } |
        Sort-Object ReceivedTime -Descending |
        Select-Object -First 1

    if (-not $email)
    {
        throw "Employee ID email not found."
    }

    $saveFolder = "C:\Temp"

    if (!(Test-Path $saveFolder))
    {
        New-Item -ItemType Directory -Path $saveFolder | Out-Null 
    }

    $Users = @()

    foreach ($attachment in $email.Attachments)
    {
        if ($attachment.FileName -match '\.(xlsx|xls)$')
        {
            $filePath = Join-Path $saveFolder $attachment.FileName
            $attachment.SaveAsFile($filePath)

            $excel = New-Object -ComObject Excel.Application
            $workbook = $excel.Workbooks.Open($filePath)
            $worksheet = $workbook.Worksheets.Item(1)

            $usedRange = $worksheet.UsedRange
            $rowCount = $usedRange.Rows.Count

            for ($row = 2; $row -le $rowCount; $row++)
            {
                $Users += [PSCustomObject]@{
                    FirstName              = $worksheet.Cells.Item($row,1).Text.Trim()
                    Surname                = $worksheet.Cells.Item($row,2).Text.Trim()
                    DateOfBirth            = $worksheet.Cells.Item($row,3).Text.Trim()
                    Gender                 = $worksheet.Cells.Item($row,4).Text.Trim()
                    ImageFile              = $worksheet.Cells.Item($row,5).Text.Trim()
                    GainsboroughHouse      = $worksheet.Cells.Item($row,6).Text.Trim()
                    FenchurchStreet        = $worksheet.Cells.Item($row,7).Text.Trim()
                    TechnicalSupportCenter = $worksheet.Cells.Item($row,8).Text.Trim()
                }
            }

            $workbook.Close($false)
            $excel.Quit()
        }
    }

    return @{
        Email = $email
        Users = $Users
    }
}