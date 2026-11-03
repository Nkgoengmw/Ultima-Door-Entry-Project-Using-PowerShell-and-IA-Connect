function Send-EmployeeResponse
{
    param(
        [Parameter(Mandatory)]
        $SourceEmail,

        [Parameter(Mandatory)]
        $Results
    )

    $reply = $SourceEmail.Reply()

    $body = "Door Entry User Creation Results`r`n`r`n"

    foreach ($result in $Results)
    {
        $body += @"

Name: $($result.FirstName) $($result.Surname)
Employee ID: $($result.EmployeeID)
Status: $($result.Status)

"@
    }

    $reply.Body = $body + "`r`n`r`n" + $reply.Body

    # Send reply
    $reply.Send()

    Write-Host "Reply sent successfully." -ForegroundColor Green

    # Move original email to Deleted Items
    try
    {
        $outlook = $SourceEmail.Application
        $namespace = $outlook.GetNamespace("MAPI")

        # 3 = Deleted Items folder
        $deletedItemsFolder = $namespace.GetDefaultFolder(3)

        $SourceEmail.Move($deletedItemsFolder)

        Write-Host "Original email moved to Deleted Items." -ForegroundColor Yellow
    }
    catch
    {
        Write-Warning "Reply was sent, but email could not be moved to Deleted Items: $($_.Exception.Message)"
    }
}