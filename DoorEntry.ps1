
#Requires -Version 5.1

$ErrorActionPreference = "Stop"

Import-Module Ultima.IAConnect.PowerShell -ErrorAction Stop

function New-DoorEntryUser
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$User,

        [Parameter(Mandatory = $false)]
        [string]$DoorEntryExe = "C:\Users\matome.nkgoeng\OneDrive - NTT Ltd\Desktop\Door entry sample application\Door entry.exe",

        [Parameter(Mandatory = $false)]
        [string]$EnvironmentName = "Testing",

        [Parameter(Mandatory = $false)]
        [string]$Username = "operator",

        [Parameter(Mandatory = $false)]
        [string]$PlainTextPassword = "*****"
    )

    # ========================================================
    # HELPER: VALIDATE IACONNECT RESULT
    # ========================================================

    function Assert-IAResult
    {
        param
        (
            [Parameter(Mandatory = $true)]
            [string]$ActionName,

            [Parameter(Mandatory = $true)]
            $Result
        )

        if ($null -eq $Result)
        {
            throw "$ActionName returned no result."
        }

        if ($Result -is [bool])
        {
            if ($Result -eq $false)
            {
                throw "$ActionName was not successful."
            }

            Write-Host "[SUCCESS] $ActionName" -ForegroundColor Green
            return
        }

        $successPropertyNames = @(
            "Result",
            "ClickElementResult",
            "PressElementResult",
            "CheckElementResult",
            "InputTextIntoElementResult",
            "InputPasswordIntoElementResult"
        )

        foreach ($propertyName in $successPropertyNames)
        {
            if ($Result.PSObject.Properties.Name -contains $propertyName)
            {
                $propertyValue = $Result.$propertyName

                if ($propertyValue -eq $false)
                {
                    $resultErrorMessage = $null

                    if ($Result.PSObject.Properties.Name -contains "ErrorMessage")
                    {
                        $resultErrorMessage = $Result.ErrorMessage
                    }

                    if ($null -ne $resultErrorMessage -and $resultErrorMessage -ne "")
                    {
                        throw "$ActionName was not successful. Error: $resultErrorMessage"
                    }

                    throw "$ActionName was not successful."
                }

                Write-Host "[SUCCESS] $ActionName" -ForegroundColor Green
                return
            }
        }

        Write-Host "[SUCCESS] $ActionName" -ForegroundColor Green
    }

    # ========================================================
    # HELPER: CONVERT EXCEL YES/NO VALUES TO BOOLEAN
    # ========================================================

    function ConvertTo-AccessBoolean
    {
        param
        (
            [Parameter(Mandatory = $false)]
            $Value
        )

        if ($Value -is [bool])
        {
            return $Value
        }

        $textValue = "$Value"
        $textValue = $textValue.Trim().ToLowerInvariant()

        switch ($textValue)
        {
            "yes"     { return $true }
            "true"    { return $true }
            "1"       { return $true }
            "y"       { return $true }
            "checked" { return $true }

            "no"        { return $false }
            "false"     { return $false }
            "0"         { return $false }
            "n"         { return $false }
            "unchecked" { return $false }
            ""          { return $false }

            default
            {
                throw "Invalid access value '$Value'. Expected Yes, No, True, False, 1, or 0."
            }
        }
    }

    # ========================================================
    # READ VALUES FROM THE USER OBJECT
    # ========================================================

    $FirstName = "$($User.FirstName)"
    $Surname = "$($User.Surname)"
    $DateOfBirth = "$($User.DateOfBirth)"
    $Gender = "$($User.Gender)"
    $ImageFile = "$($User.ImageFile)"

    $FirstName = $FirstName.Trim()
    $Surname = $Surname.Trim()
    $DateOfBirth = $DateOfBirth.Trim()
    $Gender = $Gender.Trim()
    $ImageFile = $ImageFile.Trim()

    $SelectGainsboroughHouse = ConvertTo-AccessBoolean `
        -Value $User.GainsboroughHouse

    $SelectFenchurchStreet = ConvertTo-AccessBoolean `
        -Value $User.FenchurchStreet

    $SelectTechnicalSupportCenter = ConvertTo-AccessBoolean `
        -Value $User.TechnicalSupportCenter

    $SecurePassword = $null
    $doorEntryProcess = $null
    $LoginHandle = $null
    $MainHandle = $null
    $AddUserHandle = $null
    $EmployeeID = $null

    try
    {
        Write-Host ""
        Write-Host "==========================================" -ForegroundColor Cyan
        Write-Host "Creating Door Entry user" -ForegroundColor Cyan
        Write-Host "Name: $FirstName $Surname" -ForegroundColor Cyan
        Write-Host "==========================================" -ForegroundColor Cyan

        # ====================================================
        # VALIDATE REQUIRED VALUES
        # ====================================================

        if ($FirstName -eq "")
        {
            throw "FirstName is empty."
        }

        if ($Surname -eq "")
        {
            throw "Surname is empty."
        }

        if ($DateOfBirth -eq "")
        {
            throw "DateOfBirth is empty."
        }

        if ($Gender -eq "")
        {
            throw "Gender is empty."
        }

        if ($ImageFile -eq "")
        {
            throw "ImageFile is empty."
        }

        if (-not (Test-Path -LiteralPath $DoorEntryExe -PathType Leaf))
        {
            throw "Door Entry executable was not found: $DoorEntryExe"
        }

        if (-not (Test-Path -LiteralPath $ImageFile -PathType Leaf))
        {
            throw "The image file does not exist: $ImageFile"
        }

        # ====================================================
        # 1. START DOOR ENTRY
        # ====================================================

        Write-Host "Starting Door Entry..." -ForegroundColor Cyan

        $doorEntryProcess = Start-Process `
            -FilePath $DoorEntryExe `
            -PassThru

        if ($null -eq $doorEntryProcess)
        {
            throw "Door Entry did not return a process object."
        }

        Write-Host "Door Entry process ID: $($doorEntryProcess.Id)"

        # ====================================================
        # 2. WAIT FOR LOGIN WINDOW
        # ====================================================

        Write-Host "Waiting for the login window..." -ForegroundColor Cyan

        $windowResult = Wait-IAUIAWaitForTopLevelWindow `
            -SearchWindowTitle "Door entry system login" `
            -SearchProcessId $doorEntryProcess.Id `
            -SecondsToWait 30 `
            -MatchIndex 1 `
            -RaiseExceptionIfWindowNotFound $true

        $LoginHandle = $windowResult.WindowHandle

        if (-not $LoginHandle)
        {
            throw "The Door Entry login window handle was not returned."
        }

        Write-Host "Login window handle: $LoginHandle"

        # ====================================================
        # 3. SELECT ENVIRONMENT
        # ====================================================

        $environmentResult = Use-IAUIAClickListElementItemByName `
            -ParentWindowHandle $LoginHandle `
            -SearchElementAutomationId "ComboBox_Environment" `
            -SearchLocalizedControlType "combo box" `
            -SearchSubTree $true `
            -ExpandFirst $true `
            -ItemName $EnvironmentName

        Assert-IAResult `
            -ActionName "Select environment '$EnvironmentName'" `
            -Result $environmentResult

        # ====================================================
        # 4. ENTER USERNAME
        # ====================================================

        $usernameResult = Set-IAUIAInputTextIntoElement `
            -ParentWindowHandle $LoginHandle `
            -SearchElementAutomationId "TextBox_Login_Username" `
            -SearchLocalizedControlType "edit" `
            -SearchSubTree $true `
            -TextToInput $Username `
            -ReplaceExistingValue $true

        Assert-IAResult `
            -ActionName "Enter username" `
            -Result $usernameResult

        # ====================================================
        # 5. ENTER PASSWORD
        # ====================================================

        $SecurePassword = ConvertTo-SecureString `
            -String $PlainTextPassword `
            -AsPlainText `
            -Force

        $passwordResult = Set-IAUIAInputPasswordIntoElement `
            -ParentWindowHandle $LoginHandle `
            -SearchElementAutomationId "TextBox_Login_Password" `
            -SearchLocalizedControlType "edit" `
            -SearchSubTree $true `
            -PasswordToInput $SecurePassword

        Assert-IAResult `
            -ActionName "Enter password" `
            -Result $passwordResult

        # ====================================================
        # 6. PRESS LOGIN OK
        # ====================================================

        $loginResult = Use-IAUIAPressElement `
            -ParentWindowHandle $LoginHandle `
            -SearchElementAutomationId "Button_Login_OK" `
            -SearchLocalizedControlType "button" `
            -SearchSubTree $true

        Assert-IAResult `
            -ActionName "Press login OK" `
            -Result $loginResult

        Start-Sleep -Seconds 2

        # ====================================================
        # 7. GET MAIN WINDOW HANDLE
        # ====================================================

        Write-Host "Getting the main window handle..." -ForegroundColor Cyan

        $mainWindow = Get-IAUIAGetHandleForProcessIdMainWindow `
            -ProcessId $doorEntryProcess.Id

        $MainHandle = $mainWindow.WindowHandle

        if (-not $MainHandle)
        {
            throw "The Door Entry main window handle was not returned."
        }

        Write-Host "Main window handle: $MainHandle"
        Write-Host "Main window name: $($mainWindow.ElementName)"

        # ====================================================
        # 8. PRESS ADD NEW USER
        # ====================================================

        $newUserResult = Use-IAUIAGlobalMouseClickOnElement `
            -ParentWindowHandle $MainHandle `
            -SearchElementAutomationId "Button_NewUser" `
            -SearchLocalizedControlType "button" `
            -SearchSubTree $true `
            -FocusElementFirst $true `
            -ValidateClickablePointWithinElementBoundary $true

        Assert-IAResult `
            -ActionName "Press Add New User" `
            -Result $newUserResult

        # ====================================================
        # 9. GET ADD NEW USER WINDOW HANDLE
        # ====================================================

        Write-Host "Waiting for the Add New User form..." `
            -ForegroundColor Cyan

        $addUserWindow = Wait-IAUIAWaitForElement `
            -ParentWindowHandle $MainHandle `
            -SearchElementName "Add new user" `
            -SearchElementAutomationId "AddNewUserForm" `
            -SearchLocalizedControlType "window" `
            -SearchSubTree $true `
            -SecondsToWait 30 `
            -MatchIndex 1 `
            -RaiseExceptionIfElementNotFound $true

        $AddUserHandle = $addUserWindow.ElementHandle

        if (-not $AddUserHandle)
        {
            throw "The Add New User form handle was not returned."
        }

        Write-Host "Add New User handle: $AddUserHandle"

        # ====================================================
        # 10. ENTER FIRST NAME
        # ====================================================

        $firstNameResult = Set-IAUIAInputTextIntoElement `
            -ParentWindowHandle $AddUserHandle `
            -SearchElementAutomationId "TextBox_FirstName" `
            -SearchLocalizedControlType "edit" `
            -SearchSubTree $true `
            -TextToInput $FirstName `
            -ReplaceExistingValue $true

        Assert-IAResult `
            -ActionName "Enter first name" `
            -Result $firstNameResult

        # ====================================================
        # 11. ENTER SURNAME
        # ====================================================

        $surnameResult = Set-IAUIAInputTextIntoElement `
            -ParentWindowHandle $AddUserHandle `
            -SearchElementAutomationId "TextBox_Surname" `
            -SearchLocalizedControlType "edit" `
            -SearchSubTree $true `
            -TextToInput $Surname `
            -ReplaceExistingValue $true

        Assert-IAResult `
            -ActionName "Enter surname" `
            -Result $surnameResult

        # ====================================================
        # 12. ENTER DATE OF BIRTH
        # ====================================================

        $dateOfBirthResult = Set-IAUIAInputTextIntoElement `
            -ParentWindowHandle $AddUserHandle `
            -SearchElementAutomationId "TextBox_DateOfBirth" `
            -SearchLocalizedControlType "edit" `
            -SearchSubTree $true `
            -TextToInput $DateOfBirth `
            -ReplaceExistingValue $true

        Assert-IAResult `
            -ActionName "Enter date of birth" `
            -Result $dateOfBirthResult

        # ====================================================
        # 13. ENTER GENDER
        # ====================================================

        $genderResult = Set-IAUIAInputTextIntoElement `
            -ParentWindowHandle $AddUserHandle `
            -SearchElementAutomationId "Textbox_Gender" `
            -SearchLocalizedControlType "edit" `
            -SearchSubTree $true `
            -TextToInput $Gender `
            -ReplaceExistingValue $true

        Assert-IAResult `
            -ActionName "Enter gender" `
            -Result $genderResult

        # ====================================================
        # 14. SET GAINSBOROUGH HOUSE
        # ====================================================

        $gainsboroughResult = Use-IAUIACheckElement `
            -ParentWindowHandle $AddUserHandle `
            -SearchElementAutomationId "CheckBox_GainsboroughHouse" `
            -SearchLocalizedControlType "check box" `
            -SearchSubTree $true `
            -CheckElement $SelectGainsboroughHouse

        Assert-IAResult `
            -ActionName "Set Gainsborough House access" `
            -Result $gainsboroughResult

        # ====================================================
        # 15. SET FENCHURCH STREET
        # ====================================================

        $fenchurchResult = Use-IAUIACheckElement `
            -ParentWindowHandle $AddUserHandle `
            -SearchElementAutomationId "CheckBox_FenchurchStreetOffice" `
            -SearchLocalizedControlType "check box" `
            -SearchSubTree $true `
            -CheckElement $SelectFenchurchStreet

        Assert-IAResult `
            -ActionName "Set Fenchurch Street access" `
            -Result $fenchurchResult

        # ====================================================
        # 16. SET TECHNICAL SUPPORT CENTER
        # ====================================================

        $tscResult = Use-IAUIACheckElement `
            -ParentWindowHandle $AddUserHandle `
            -SearchElementAutomationId "CheckBox_TSC" `
            -SearchLocalizedControlType "check box" `
            -SearchSubTree $true `
            -CheckElement $SelectTechnicalSupportCenter

        Assert-IAResult `
            -ActionName "Set Technical Support Center access" `
            -Result $tscResult

        # ====================================================
        # 17. ENTER IMAGE FILE
        # ====================================================

        $imageResult = Set-IAUIAInputTextIntoElement `
            -ParentWindowHandle $AddUserHandle `
            -SearchElementAutomationId "TextBox_ImageFilename" `
            -SearchLocalizedControlType "edit" `
            -SearchSubTree $true `
            -TextToInput $ImageFile `
            -ReplaceExistingValue $true

        Assert-IAResult `
            -ActionName "Enter image filename" `
            -Result $imageResult

        # ====================================================
        # 18. PRESS SUBMIT
        # ====================================================

        $submitResult = Use-IAUIAGlobalMouseClickOnElement `
            -ParentWindowHandle $AddUserHandle `
            -SearchElementName "Submit" `
            -SearchElementAutomationId "Button_Submit" `
            -SearchLocalizedControlType "button" `
            -SearchSubTree $true `
            -FocusElementFirst $true `
            -ValidateClickablePointWithinElementBoundary $true

        Assert-IAResult `
            -ActionName "Press Submit" `
            -Result $submitResult

        Start-Sleep -Seconds 2

        # ====================================================
        # 19. FIND AND EXTRACT EMPLOYEE ID
        # ====================================================

        Write-Host "Searching for the Employee ID..." `
            -ForegroundColor Cyan

        $elements = Get-IAUIAGetMultipleElementProperties `
            -ParentWindowHandle $MainHandle `
            -SearchDescendants $true `
            -ReturnElementHandle $true `
            -ReturnElementValue $true

        if ($null -eq $elements)
        {
            throw "No UI elements were returned after submitting the form."
        }

        if ($null -eq $elements.ElementProperties)
        {
            throw "No element properties were returned after submitting the form."
        }

        $messageElement = $elements.ElementProperties |
            Where-Object {
                "$($_.Name)" -match "(?i)employee\s*ID\s*:\s*\d+"
            } |
            Select-Object -First 1

        if ($null -eq $messageElement)
        {
            throw "The Employee ID message was not found."
        }

        $message = "$($messageElement.Name)"
        $message = $message.Trim()

        if ($message -match "(?i)employee\s*ID\s*:\s*(\d+)")
        {
            $EmployeeID = [int]$Matches[1]
        }
        else
        {
            throw "Could not extract the Employee ID from: $message"
        }

        Write-Host "Popup message: $message"
        Write-Host "Employee ID: $EmployeeID" -ForegroundColor Green

        # ====================================================
        # 20. CLOSE ENTRY SUBMITTED POPUP
        # ====================================================

        $popupOkResult = Use-IAUIAPressElement `
            -ParentWindowHandle $MainHandle `
            -SearchElementName "OK" `
            -SearchElementAutomationId "2" `
            -SearchLocalizedControlType "button" `
            -SearchSubTree $true

        Assert-IAResult `
            -ActionName "Close Entry Submitted popup" `
            -Result $popupOkResult

        # ====================================================
        # 21. RETURN SUCCESS RESULT
        # ====================================================

        Write-Host ""
        Write-Host "Door Entry automation completed successfully." `
            -ForegroundColor Green

        return [PSCustomObject]@{
            Environment            = $EnvironmentName
            Username               = $Username
            FirstName              = $FirstName
            Surname                = $Surname
            DateOfBirth            = $DateOfBirth
            Gender                 = $Gender
            GainsboroughHouse      = $SelectGainsboroughHouse
            FenchurchStreet        = $SelectFenchurchStreet
            TechnicalSupportCenter = $SelectTechnicalSupportCenter
            ImageFile              = $ImageFile
            EmployeeID             = $EmployeeID
            Status                 = "Success"
            ErrorMessage           = $null
        }
    }
    catch
    {
        $errorMessage = $_.Exception.Message

        Write-Host ""
        Write-Host "DOOR ENTRY ERROR" -ForegroundColor Red
        Write-Host $errorMessage -ForegroundColor Red

        if ($null -ne $_.ScriptStackTrace)
        {
            Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
        }

        return [PSCustomObject]@{
            Environment            = $EnvironmentName
            Username               = $Username
            FirstName              = $FirstName
            Surname                = $Surname
            DateOfBirth            = $DateOfBirth
            Gender                 = $Gender
            GainsboroughHouse      = $SelectGainsboroughHouse
            FenchurchStreet        = $SelectFenchurchStreet
            TechnicalSupportCenter = $SelectTechnicalSupportCenter
            ImageFile              = $ImageFile
            EmployeeID             = $null
            Status                 = "Failed"
            ErrorMessage           = $errorMessage
        }
    }
    finally
    {
        $SecurePassword = $null
    }
}