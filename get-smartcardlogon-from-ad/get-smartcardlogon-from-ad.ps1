<#
.SYNOPSIS
Queries Active Directory for users required to use a smartcard for logon.

.DESCRIPTION
Designed to query Active Directory for users who have the smartcard logon requirement configured. Provides details on the smartcard logon configurations for further analysis.

.EXAMPLE
.\get-smartcardlogon-from-ad.ps1 -ou "OU=ouname,DC=test,DC=local"

Executes the script to retrieve and report on users with smartcard logon configurations in Active Directory.

.PARAMETER CustomParam
Describe any parameters that the script might take.

.NOTES
Additional information about the script.

#>
[CmdletBinding(PositionalBinding = $False)]
param (
    [Alias("ou")]
    [Parameter(Mandatory = $False, HelpMessage = "Enter an Organizational Unit (System.Array).")]
    [ValidatePattern("^(OU|DC)=\w+[\s,\.\w=]*$")]
    [Array]$organizational_units
)
Function GetSmartCardLogonInfo {
    param (
        [Parameter(Mandatory = $true)]
        [Array]$organizational_units,
        [Parameter(Mandatory = $true)]
        [Boolean]$NonInteractive
    )
    try {
    # Starting Excel
    $Excel = New-Object -ComObject Excel.Application 
    # Excel window visibility
    $Excel.Visible = $false
    # Adding a workbook
    $ExcelWorkBook = $Excel.Workbooks.Add()
    # Connecting to the first sheet
    $ExcelWorkSheet = $ExcelWorkBook.Worksheets.Item(1)
    # Renaming the sheet
    $ExcelWorkSheet.Name = 'Отчет по смарт-картам'
    # Adding a header
    $ExcelWorkSheet.Cells.Item(1,1) = 'Перечень УЗ, у которых отключен параметр "Вход только по смарт-карте":'
    # Font settings in the header
    $ExcelWorkSheet.Rows.Item(1).Font.Bold = $true
    $ExcelWorkSheet.Rows.Item(1).Font.size= 14
    # Center aligning
    $Range = $ExcelWorkSheet.Range('A1','K1')
    $Range.Merge()
    # Vertical aligning
    $Range.VerticalAlignment = -4108
    $Range.HorizontalAlignment = -4108
    # Adding column names
    $ExcelWorkSheet.Cells.Item(2,1) = "№ п/п"
    $ExcelWorkSheet.Cells.Item(2,2) = "ФИО"
    $ExcelWorkSheet.Cells.Item(2,3) = "Имя УЗ"
    $ExcelWorkSheet.Cells.Item(2,4) = "Статус УЗ"
    $ExcelWorkSheet.Cells.Item(2,5) = "Срок действия пароля истек"
    $ExcelWorkSheet.Cells.Item(2,6) = "Дата последнего изменения пароля"
    $ExcelWorkSheet.Cells.Item(2,7) = "Срок действия пароля не ограничен"
    $ExcelWorkSheet.Cells.Item(2,8)= "Вход только по смарт-карте"
    $ExcelWorkSheet.Cells.Item(2,9)= "Практиканты/`nПодрядчики"
    $ExcelWorkSheet.Cells.Item(2,10)= "Размещение"
    $ExcelWorkSheet.Cells.Item(2,11)= "Дата создания УЗ"
    # Font settings in the title
    $ExcelWorkSheet.Rows.Item(2).Font.Bold = $true
    $ExcelWorkSheet.Rows.Item(2).HorizontalAlignment = -4108
    # Shading the first two rows in gray
    $ExcelWorkSheet.Range('A1','K1').Interior.Color = 14277081
    $ExcelWorkSheet.Range('A2','K2').Interior.Color = 14277081
    # Setting the autofilter in the second row
    $headerRange = $ExcelWorkSheet.Range('A2', 'K2')
    $headerRange.AutoFilter() | Out-Null
    # Moving to the 3rd row
    $Row = 3
    # Sequential row number in the table
    $i = 1
    # Counter for the progress bar
    $counter = 1
    # Retrieving computers in OU
    $adUsers = foreach ($ou in $organizational_units) {
         Get-ADUser -SearchBase $OU -Filter * -properties SamAccountName, PasswordExpired, PasswordLastSet, PasswordNeverExpires, SmartcardLogonRequired, StreetAddress, Company, Created | Select-Object Name, SamAccountName, Enabled, PasswordExpired, PasswordLastSet, PasswordNeverExpires, SmartcardLogonRequired, StreetAddress, Company, Created | Sort-Object Name
    }
    write-host "Всего учетных записей: $($adUsers.count)"
    # Iterating through AD computers
    foreach ($adUser in $adUsers) {
        # Progress bar
        write-progress -Activity 'Прогресс выполнения:' -PercentComplete ($counter/$adUsers.Count*100)
        # If login with smart-card is disabled
        if ((($aduser.Enabled) -and (!$aduser.SmartcardLogonRequired)) -or (($aduser.PasswordNeverExpires) -and ($aduser.Enabled))) {
            # Sequential row number
            $ExcelWorkSheet.Cells.Item($Row,1) = [string]$i 
            # Full Name
            $ExcelWorkSheet.Cells.Item($Row,2) = [string]$aduser.Name
            # Account Name
            $ExcelWorkSheet.Cells.Item($Row,3) = ([string]$aduser.SamAccountName).ToLower()
            # Account Status
            if ($aduser.Enabled) {
                $ExcelWorkSheet.Cells.Item($Row,4) = 'Активна'
            }
            else {
                $ExcelWorkSheet.Cells.Item($Row,4) = 'Отключена'
            }
            # Password expiration date has passed
            if ($aduser.PasswordExpired) {
                $ExcelWorkSheet.Cells.Item($Row,5) = 'Да'
                # Highlighting the cell in yellow
                $ExcelWorkSheet.Cells.Item($Row,5).Interior.ColorIndex = 6
            }
            else {
                $ExcelWorkSheet.Cells.Item($Row,5) = 'Нет'
            }
            # Date of last password change
            if ($aduser.PasswordLastSet) {
                $ExcelWorkSheet.Cells.Item($Row,6) = ($aduser.PasswordLastSet).Tostring('dd.MM.yyyy HH:MM')
            }
            # Password expiration is not limited
            if ($aduser.PasswordNeverExpires) {
                $ExcelWorkSheet.Cells.Item($Row,7) = 'Да'
                # Highlighting the cell in yellow
                $ExcelWorkSheet.Cells.Item($Row,7).Interior.ColorIndex = 6
            }
            else {
                $ExcelWorkSheet.Cells.Item($Row,7) = 'Нет'
            }
            # Login with smart-card only
            if ($aduser.SmartcardLogonRequired) {
                $ExcelWorkSheet.Cells.Item($Row,8) = 'Да'
            }
            else {
                $ExcelWorkSheet.Cells.Item($Row,8) = 'Нет'
                # Highlighting the cell in yellow
                $ExcelWorkSheet.Cells.Item($Row,8).Interior.ColorIndex = 6
            }
            # Interns/Contractors
            if (($aduser.Company).ToLower() -in @('подрядчики', 'практиканты')) {
                $ExcelWorkSheet.Cells.Item($Row,9) = [string]$aduser.Company
            }
            # Placement
            $ExcelWorkSheet.Cells.Item($Row,10) = [string]$aduser.StreetAddress
            # Creation date
            $ExcelWorkSheet.Cells.Item($Row,11) = [string]$aduser.Created
            # Incrementing counters
            $i++
            $counter++
            $Row++
        }
        else {
            $counter++
        }
    }
    # Moving one row up
    $Row--
    # Getting information about the range of cells filled in the table
    $TableRange = $ExcelWorkSheet.Range('A1', ('K{0}' -f $Row--))
    # Drawing borders around cells in the filled table
    $xlcontinuous = 1
    $TableRange.Borders.LineStyle = $xlcontinuous
    $UsedRange = $ExcelWorkSheet.Range('A2', 'K2')
    $UsedRange.EntireColumn.AutoFit() | Out-Null
    $UsedRange.WrapText = $true
    $ExcelWorkSheet.columns.item('A').ColumnWidth = 10
    $ExcelWorkSheet.columns.item('E').ColumnWidth = 18
    $ExcelWorkSheet.columns.item('F').ColumnWidth = 22
    $ExcelWorkSheet.columns.item('G').ColumnWidth = 22
    $ExcelWorkSheet.columns.item('H').ColumnWidth = 20
    $ExcelWorkSheet.columns.item('I').ColumnWidth = 17
    $TableRange.VerticalAlignment = -4160
    # Path and name of the file to be saved
    $File = $PSScriptRoot + "\Отчет по смарт-картам_" + (Get-Date -Format 'dd/MM/yyyy HH mm ss') +  '.xlsx'
    # Saving the file and closing Excel
    $ExcelWorkBook.SaveAs($File)
    $ExcelWorkBook.close($true)
    $Excel.Quit()
    if (!$NonInteractive){
        # Displaying a message about successful completion
        [System.Windows.Forms.MessageBox]::Show("Отчет сформирован!`nВсего учетных записей: $($adUsers.count)","Готово",[System.Windows.Forms.MessageBoxButtons]::OK)
    }
    else { Write-Host $File }
    }
    # If errors occurred
    catch {
        if (!$NonInteractive) {
            # Displaying an error message
            [System.Windows.Forms.MessageBox]::Show("При выполнении программы возникли ошибки: `n $($Error[0])","Внимание",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        exit
    }
}

# Determining the script execution mode
$NonInteractive = @($False, $True)[$PSBoundParameters.Count -ge 1]

# Passing a list of OUs to scan and the OS build number to the function
GetSmartCardLogonInfo -organizational_units $organizational_units -NonInteractive $NonInteractive