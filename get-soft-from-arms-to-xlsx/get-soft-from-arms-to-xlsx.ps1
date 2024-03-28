<#
.SYNOPSIS
Exports a list of software from ARMS to an Excel file.

.DESCRIPTION
Interfaces with the Asset Management System (ARMS) to retrieve a list of software installed on systems. Exports the gathered data to an Excel (.xlsx) file for analysis and reporting.

.EXAMPLE
.\get-soft-from-arms-to-xlsx.ps1 -ou "OU=ouname,DC=test,DC=local" -d "C:\Your\Directory\Path" -t 15 -c 2 a "$False"

This example demonstrates how to run the script and export the software list to a specified Excel file.

.PARAMETER Output
The output path for the Excel file to be saved.

.NOTES
Additional information about the script.

#>
[CmdletBinding(PositionalBinding = $False)]
param (
    [Alias("ou")]
    [Parameter(Mandatory = $False, HelpMessage = "Enter an Organizational Unit (System.Array).")]
    [ValidatePattern("^(OU|DC)=\w+[\s,\.\w=]*$")]
    [Array]$OUs,
    
    [Alias("d")]
    [Parameter(Mandatory = $False, HelpMessage = "Json file upload directory.")]
    [String]$dir,
    
    [Alias("t")]
    [Parameter(Mandatory = $False, HelpMessage = "Еhe number of workstations from which the software list is simultaneously unloaded.")]
    [ValidateScript({
			if ([Int32]$_ -ge 1) { $true }
			else { Throw "The number of threads must be [Int32] and -ge 1." }
	})]
    [Int32]$maxThreads = 15,
    
    [Alias("c")]
    [Parameter(Mandatory = $False, HelpMessage = "Number of requests sent for Test-Connection (1-2).")]
    [ValidateScript({
			if ([Int32]$_ -ge 1) { $true }
			else { Throw "The number of packets must be [Int32] and -ge 1." }
	})]
    [Int32]$packetCount = 2,
    
    [Alias("a")]
    [Parameter(Mandatory = $False, HelpMessage = "The parameter adds the names of the APMs to the report.")]
    [ValidateSet($False, $True)]
    [Boolean]$armNames = $True
)

# Software export to JSON from a separate workstation
$GetSoftFromArms = {

    param(
        [Parameter(Mandatory = $True)]
        [String]$computer,
        [Parameter(Mandatory = $True)]
        [String]$dir,
        [Parameter(Mandatory = $True)]
        [Int32]$packetCount
    )

    # Checking workstation availability (ping) and converting to bool
    $test = Test-Connection -ComputerName $computer -Count $packetCount -Quiet

    # If the workstation is available
    if($test) {
        try
	    { 
            
            # Changing the startup mode of the WINRM service
            set-service -name WINRM -ComputerName $computer -StartupType Auto -ErrorAction Stop
            # Starting the WINRM service
            get-service -name WINRM -ComputerName $computer | start-service -ErrorAction Stop
	    }
	    catch
	    {
            return
        }

        try
	    {
            # Creating a remote session
            $session = New-PSSession -ComputerName $computer -ErrorAction Stop # Remote session

            # List of installed software, excluding system software (metro apps)
            $software = Invoke-Command -Session $session -ScriptBlock {
		
		
		        $OSArchitecture = (gwmi win32_operatingSystem -ErrorAction Stop).OSArchitecture # Checking the architecture (64 or 86)
                # 64-bit architecture
		        if ($OSArchitecture -like '*64*')
		        {
			        if (Test-Path -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\')
			        {
				        $list1 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion | Where-Object { $_.DisplayName -notlike '*Security Update*' } # Without OS updates
			        }
			        if (Test-Path -Path 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\')
			        {
				        $list2 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion | Where-Object { $_.DisplayName -notlike '*Security Update*' } # Without OS updates
			        }
			        if (Test-Path -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\')
			        {
				        $list3 = Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion | Where-Object { $_.DisplayName -notlike '*Security Update*' } # Without OS updates
			        }
			
			        $list1 + $list2 + $list3 | Where-Object { $_.DisplayName } | Sort-Object DisplayName | Get-Unique -AsString # Removed empty lines and duplicates
		        }
		        # 86-bit architecture
		        else
		        {
			        if (Test-Path -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\')
			        {
				        $list1 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion | Where-Object { $_.DisplayName -notlike '*Security Update*' } # Without OS updates
			        }
			        if (Test-Path -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\')
			        {
				        $list2 = Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion | Where-Object { $_.DisplayName -notlike '*Security Update*' } # Without OS updates
			        }
			        $list1 + $list2 | Where-Object { $_.DisplayName } | Sort-Object DisplayName | Get-Unique -AsString # Removed empty lines and duplicates
		        }
	        }
        }
        catch
	    {
            # Deleting the session
            if ($session) {
                Remove-PSSession $session
            }
            return
        }
        # Deleting the session
        if ($session) {
                Remove-PSSession $session
        }
        # JSON unload path
        $path = "$($dir)$($computer).json"
        $software | Select-Object DisplayName, DisplayVersion | ConvertTo-Json >> $path
        return
    }
    # If the workstation is unavailable
    else {
        return
    }
}

# Exporting software from the workstation to json
function ExportSoftToJson {
    
    param(
        [Parameter(Mandatory = $True)]
        [Object]$OUs,
        [Parameter(Mandatory = $True)]
        [Object]$ScriptBlock,
        [Parameter(Mandatory = $True)]
        [Int32]$packetCount,
        [Parameter(Mandatory = $True)]
        [Int32]$maxThreads,
        [Parameter(Mandatory = $True)]
        [String]$dir,
        [Parameter(Mandatory = $True)]
        [Boolean]$armNames,
        [Parameter(Mandatory = $true)]
        [Boolean]$NonInteractive
    )

    # Multithreading
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $maxThreads)
    $RunspacePool.Open()
    $Jobs = @()
    # Creating a directory for json export
    New-Item -Path $dir -ItemType Directory | Out-Null
    # Exporting the list of workstations
    $adComputers = foreach ($OU in $OUs) {
        Get-ADComputer -SearchBase $OU -Filter * -properties Name| Select-Object Name | Sort-Object Name
    }
    # Iterating through workstations
    foreach ($adComputer in $adComputers) {
	    $PowerShell = [powershell]::Create() 
	    $PowerShell.RunspacePool = $RunspacePool
        # Export
	    $PowerShell.AddScript($ScriptBlock).AddArgument($adComputer.Name).AddArgument($dir).AddArgument($packetCount) 
	    $Jobs += $PowerShell.BeginInvoke() 
    }

    # Monitoring task execution
    while ($Jobs.IsCompleted -contains $false) {
	    Start-Sleep -Milliseconds 100
        # Number of tasks completed
        $count = ($Jobs.IsCompleted | Group-Object | Where-Object { $_.Name -eq $true}).count
        # Progressbar
        write-progress -Activity 'Выгрузка json c АРМ' -PercentComplete ($count/($Jobs.Count + 1)*100)
    }
    $RunspacePool.Close()
    # Comprehensive list of software from all workstations
    $dict_soft = @{}
    # List of exported jsons
    $files = Get-ChildItem -Path "$($dir)*.json" -Recurse -Force
    Write-Host "Всего выгружено $($files.Count) JSON-файлов"
    # Iterating through files
    foreach ($file in $files) {
        #Получаем содержимое json
        $software = Get-Content $file.FullName | ConvertFrom-Json
        #Имя АРМ
        $computer = ($file.Name).Replace('.json', '')
        # Retrieving the contents of json
        foreach ($app in $software) {
                # Workstation name
                $disp = [string]$app.DisplayName
                # Iterating through applications
                $vers = [string]$app.DisplayVersion
                # Software name
		        if ($disp -notin $dict_soft.Keys)
		        {
                    # Software version
                    if ($armNames) { $dict_soft[$disp] = @(@($vers), @($computer))  }
                     # If the software is not in the comprehensive list
                    else {$dict_soft[$disp] = @($vers)}  
		        }
                # Adding name, version, and workstation name
		        else
		        {
                    if ($armNames) { 
                        # Adding name and version
                        $dict_soft[$disp][1] += $computer 
                        # If the software is in the comprehensive list
                        if ($vers -notin $dict_soft[$disp][0]) {
                            # Adding the workstation name
                            $dict_soft[$disp][0] += $vers
                        }
			        }
                    else {
                        # Checking if the software version is in the comprehensive list
                        if ($vers -notin $dict_soft[$disp]) {
                            # Adding the version
                            $dict_soft[$disp] += $vers
                        }
                    }
                }
	    }
    }
    # Initiating report generation
    CreateReport -apps $dict_soft -armNames $armNames -NonInteractive $NonInteractive
    # Exporting the entire list of software to a JSON file
    $dict_soft | ConvertTo-Json  > "$PSScriptRoot\all_soft.json"
    # Deleting the directory where json files are exported
    Remove-Item -Path $dir -Recurse -Force 
}

# Generating a report in Excel format
function CreateReport {
    
    param(
        [Parameter(Mandatory = $true)]
        [Object]$apps,
        [Parameter(Mandatory = $true)]
        [Boolean]$armNames,
        [Parameter(Mandatory = $true)]
        [Boolean]$NonInteractive
    )
    try {
            write-progress -Activity 'Формирование отчета' -PercentComplete 99
            # Starting Excel
            $Excel = New-Object -ComObject Excel.Application 
            # Excel window visibility
            $Excel.Visible = $false
            # Adding a workbook
            $ExcelWorkBook = $Excel.Workbooks.Add()
            # Connecting to the first sheet
            $ExcelWorkSheet = $ExcelWorkBook.Worksheets.Item(1)
            # Renaming the sheet
            $ExcelWorkSheet.Name = 'ПО на АРМ'
            # Adding the header
            $ExcelWorkSheet.Cells.Item(1,1) = 'ПО на АРМ'
            # Setting the column C type to text
            $ExcelWorkSheet.columns.item('C').NumberFormat = "@"
            # Text alignment
            $ExcelWorkSheet.columns.item('A').VerticalAlignment = -4108
            $ExcelWorkSheet.columns.item('B').VerticalAlignment = -4108
            $ExcelWorkSheet.columns.item('C').VerticalAlignment = -4160
            if ($armNames) { $ExcelWorkSheet.columns.item('D').VerticalAlignment = -4160 }
            # Font settings in the header
            $ExcelWorkSheet.Rows.Item(1).Font.Bold = $true
            $ExcelWorkSheet.Rows.Item(1).Font.size= 14
            # Center aligning
            if ($armNames) {
                $Range = $ExcelWorkSheet.Range('A1','D1')
            } else { $Range = $ExcelWorkSheet.Range('A1','C1') } 
            $Range.Merge()
            # Vertical aligning
            $Range.VerticalAlignment = -4108
            $Range.HorizontalAlignment = -4108
            # Adding column names
            $ExcelWorkSheet.Cells.Item(2,1) = "№ п/п"
            $ExcelWorkSheet.Cells.Item(2,2) = "Название"
            $ExcelWorkSheet.Cells.Item(2,3) = "Версии"
            # Font settings in the title
            $ExcelWorkSheet.Rows.Item(2).Font.Bold = $true
            $ExcelWorkSheet.Rows.Item(2).HorizontalAlignment = -4108
            if ($armNames) {
                # Adding a column name
                $ExcelWorkSheet.Cells.Item(2,4) = "АРМы"
                # Setting the autofilter in the second row
                $headerRange = $ExcelWorkSheet.Range('A2', 'D2')
                # Shading the first row in gray
                $ExcelWorkSheet.Range('A1','D1').Interior.Color = 14277081
                $ExcelWorkSheet.Range('A2','D2').Interior.Color = 14277081 
            }
            else {
                # Setting the autofilter in the second row
                $headerRange = $ExcelWorkSheet.Range('A2', 'C2')
                $ExcelWorkSheet.Range('A1','C1').Interior.Color = 14277081
                $ExcelWorkSheet.Range('A2','C2').Interior.Color = 14277081 
            }
            $headerRange.AutoFilter() | Out-Null
            # Moving to the 3rd row
            $Row = 3
            # Sequential row number in the table
            $counter = 1
            # Iterating through applications
            foreach ($app in $apps.Keys | Sort-Object) {
                # Excluding updates
                if($app -notmatch 'Update for') {
                    # Sequential row number
                    $ExcelWorkSheet.Cells.Item($Row,1) = [string]$counter
                    # Software name
                    $ExcelWorkSheet.Cells.Item($Row,2) = [string]$app
                    # Versions
                    $ExcelWorkSheet.Cells.Item($Row,3) = ([string]($apps[$app][0] -join ",`n")).Trim()
                    if ($armNames) {
                        # Workstations
                        $ExcelWorkSheet.Cells.Item($Row,4) = ([string]($apps[$app][1] -join ", ")).Trim()
                    }
                    # Incrementing counters
                    $counter++
                    $Row++
                }  
            }
            # Moving one row up
            $Row--
            if ($armNames) {
                # Getting information about the cell range filled in the table
                $TableRange = $ExcelWorkSheet.Range('A1', ('D{0}' -f $Row--))
                # Setting the "Fit text in cell" property
                $ExcelWorkSheet.columns.item('D').WrapText = $true
                # Setting column width
                $ExcelWorkSheet.columns.item('D').columnWidth = 100
            } else { $TableRange = $ExcelWorkSheet.Range('A1', ('C{0}' -f $Row--)) }
            # Drawing cell borders in the filled table
            $xlcontinuous = 1
            $TableRange.Borders.LineStyle = $xlcontinuous
            # Aligning columns by width
            $UsedRange = $ExcelWorkSheet.UsedRange
            $UsedRange.EntireColumn.AutoFit() | Out-Null
            # Setting column width
            $ExcelWorkSheet.columns.item('A').ColumnWidth = 10
            # Path and name of the file to be saved
            $File = $PSScriptRoot + "\Софт АРМ_" + (Get-Date -Format 'dd/MM/yyyy HH mm ss') +  '.xlsx'
            # Saving the file and closing Excel
            $ExcelWorkBook.SaveAs($File)
            $ExcelWorkBook.close($true)
            $Excel.Quit()
            write-progress -Activity 'Формирование отчета:' -PercentComplete 100
            if (!$NonInteractive){
                # Displaying a message about successful completion
                [System.Windows.Forms.MessageBox]::Show("Отчет сформирован!","Готово",[System.Windows.Forms.MessageBoxButtons]::OK)
            } else { Write-Host $File } 
        }
    catch {
        if (!$NonInteractive){
            # Displaying an error message
            [System.Windows.Forms.MessageBox]::Show("В процессе выполнения произошла ошибка`n $($Error[0])","Внимание",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        $ExcelWorkBook.close($true)
        $Excel.Quit()
        return
    }
}
### Main ###
# Determining the script execution mode and performing the function
$NonInteractive = @($False, $True)[$PSBoundParameters.Count -ge 1]
if (!$PSBoundParameters.dir) {
    # Default directory for json export
    $dir = $PSScriptRoot + '\' + (Get-Date -Format 'dd/MM/yyyy HH mm ss') + '\'
} else { $dir = $PSBoundParameters.dir }
# Json export and report formation
ExportSoftToJson -OUs $OUs -ScriptBlock $GetSoftFromArms -packetCount $packetCount -MaxThreads $maxThreads -dir $dir -armNames $armNames -NonInteractive $NonInteractive | Out-Null