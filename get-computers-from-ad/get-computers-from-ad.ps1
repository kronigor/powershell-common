<#
.SYNOPSIS
Fetches a list of computer objects from Active Directory.

.DESCRIPTION
This script retrieves computer objects from Active Directory. It can be customized to filter and retrieve specific attributes of the computer objects.

.EXAMPLE
.\get-computers-from-ad.ps1 -ou "OU=ouname,DC=test,DC=local" -bn "10705"

This example runs the script without any parameters, retrieving the default set of computer object attributes from Active Directory.

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
    [Array]$organizational_units,
    [Alias("bn")]
	[Parameter(Mandatory = $False, HelpMessage = "Enter an Operating System build number (System.String).")]
	[ValidatePattern("^[0-9]+$")]
	[String]$build_number
)

Function GetADComputersInfo {
    param (
        [Parameter(Mandatory = $true)]
        [Array]$organizational_units,
        [Parameter(Mandatory = $true)]
        [String]$build_number,
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
        $ExcelWorkSheet.Name = 'Перечень АРМ'
        # Adding a header
        $ExcelWorkSheet.Cells.Item(1,1) = 'Перечень АРМ'
        # Font settings in the header
        $ExcelWorkSheet.Rows.Item(1).Font.Bold = $true
        $ExcelWorkSheet.Rows.Item(1).Font.size= 14
        # Center aligning
        $Range = $ExcelWorkSheet.Range('A1','G1')
        $Range.Merge()
        # Vertical aligning
        $Range.VerticalAlignment = -4108
        $Range.HorizontalAlignment = -4108
        # Adding column names
        $ExcelWorkSheet.Cells.Item(2,1) = "№ п/п"
        $ExcelWorkSheet.Cells.Item(2,2) = "Имя АРМ"
        $ExcelWorkSheet.Cells.Item(2,3) = "IP адрес"
        $ExcelWorkSheet.Cells.Item(2,4) = "ОС"
        $ExcelWorkSheet.Cells.Item(2,5) = "Версия ОС"
        $ExcelWorkSheet.Cells.Item(2,6) = "Дата создания"
        $ExcelWorkSheet.Cells.Item(2,7) = "Дата изменения"
        # Font settings in the title
        $ExcelWorkSheet.Rows.Item(2).Font.Bold = $true
        $ExcelWorkSheet.Rows.Item(2).HorizontalAlignment = -4108
        # Shading the first two rows in gray
        $ExcelWorkSheet.Range('A1','G1').Interior.Color = 14277081
        $ExcelWorkSheet.Range('A2','G2').Interior.Color = 14277081
        # Setting the autofilter in the second row
        $headerRange = $ExcelWorkSheet.Range('A2', 'G2')
        $headerRange.AutoFilter() | Out-Null
        # Moving to the 3rd row
        $Row = 3
        # Sequential row number in the table
        $counter = 1
        # Retrieving computers in OU
        $adComputers = foreach ($ou in $organizational_units) {
            Get-ADComputer -SearchBase $ou -Filter * -Properties CN, IPv4Address, OperatingSystem, OperatingSystemVersion, Modified, Created | Select-Object CN, IPv4Address, OperatingSystem, OperatingSystemVersion, Modified, Created | Sort-Object CN
        }
        # Iterating through AD computers
        foreach ($adComputer in $adComputers) {
                # Progress bar
                write-progress -Activity 'Прогресс выполнения:' -PercentComplete ($counter/$adComputers.Count*100)
                # Sequential row number
                $ExcelWorkSheet.Cells.Item($Row,1) = [string]$counter 
                # Workstation name
                $ExcelWorkSheet.Cells.Item($Row,2) = [string]$adComputer.CN
                # IP address
                $ExcelWorkSheet.Cells.Item($Row,3) = [string]$adComputer.IPv4Address
                # OS
                $ExcelWorkSheet.Cells.Item($Row,4) = [string]$adComputer.OperatingSystem            
                # OS version
                $ExcelWorkSheet.Cells.Item($Row,5) = [string]$adComputer.OperatingSystemVersion
                # If the OS version does not match the current version
                if ([string]$adComputer.OperatingSystemVersion -notmatch $build_number) {
                    # Highlighting the cell in yellow
                    $ExcelWorkSheet.Cells.Item($Row,5).Interior.ColorIndex = 6
                }
                # Creation date
                $ExcelWorkSheet.Cells.Item($Row,6) = [string](Get-Date($adComputer.Created) -Format 'dd.MM.yyyy HH:mm:ss')
                # Modification date
                $ExcelWorkSheet.Cells.Item($Row,7) = [string](Get-Date($adComputer.Modified) -Format 'dd.MM.yyyy HH:mm:ss' )
                # Incrementing counters
                $counter++
                $Row++
            }
        # Moving one row up
        $Row--
        # Getting information about the range of cells filled in the table
        $TableRange = $ExcelWorkSheet.Range('A1', ('G{0}' -f $Row--))
        # Drawing borders around cells in the filled table
        $xlcontinuous = 1
        $TableRange.Borders.LineStyle = $xlcontinuous
        # Aligning columns by width
        $UsedRange = $ExcelWorkSheet.UsedRange
        $UsedRange.EntireColumn.AutoFit() | Out-Null
        # Path and name of the file to be saved
        $File = $PSScriptRoot + "\Отчет по АРМ_" + (Get-Date -Format 'dd/MM/yyyy HH mm ss') +  '.xlsx'
        # Saving the file and closing Excel
        $ExcelWorkBook.SaveAs($File)
        $ExcelWorkBook.close($true)
        $Excel.Quit()
        if (!$NonInteractive){
            # Displaying a message about successful completion
            [System.Windows.Forms.MessageBox]::Show("Отчет сформирован!","Готово",[System.Windows.Forms.MessageBoxButtons]::OK)
        }
        else {
            Write-Host $File
        }
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

# Passing a list of OUs for scanning and the OS build number to the function
GetADComputersInfo -organizational_units $organizational_units -build_number $build_number -NonInteractive $NonInteractive