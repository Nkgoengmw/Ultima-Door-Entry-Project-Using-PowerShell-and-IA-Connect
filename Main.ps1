Clear-Host


."$PSScriptRoot\ReadEmail.ps1"
."$PSScriptRoot\DoorEntry.ps1"
."$PSScriptRoot\SendResponse.ps1"

try
{
    Write-Host "Reading email..."

    $mailData = Get-UsersFromEmail

    $Users = $mailData.Users
    $SourceEmail = $mailData.Email

    Write-Host "$($Users.Count) users found."

    $Results = @()

    foreach ($User in $Users)
    {
        Write-Host ""
        Write-Host "Processing $($User.FirstName) $($User.Surname)"

        $Result = New-DoorEntryUser -User $User

        $Results += $Result
    }

   Write-Host "SourceEmail Object:"
    $SourceEmail | Format-List Subject,SenderName,SenderEmailAddress
 
    Write-Host "Results Count: $($Results.Count)"

    Send-EmployeeResponse -SourceEmail $SourceEmail -Results $Results

    Write-Host ""
    Write-Host "Process completed."
    Get-Process "Door entry" -ErrorAction SilentlyContinue | Stop-Process -Force
}
catch
{
    Write-Error $_.Exception.Message
}