<#
.SYNOPSIS
Generates a detailed report of files in a specified directory.

.DESCRIPTION
This script generates a report including information such as file size, creation date, and last access date for all files within a specified directory. Useful for inventory and monitoring purposes.

.EXAMPLE
.\get-filesinfo-from-directory.ps1 -p "C:\Your\Directory\Path" -mode "hash"

Runs the script for the specified directory path, generating a report on the files contained within.

.PARAMETER Path
The path of the directory to scan for files.

.NOTES
Additional information about the script.

#>
[CmdletBinding(PositionalBinding = $False)]
param (
    [Alias("p")]
    [Parameter(Mandatory = $False, HelpMessage = "Enter the path of the directory.")]
	[ValidateScript({
			if (Test-Path $_) { $true }
			else { Throw "Path doesn't exist!" }
	})]
    [String]$path,
    [Alias("m")]
	[Parameter(Mandatory = $False, HelpMessage = "Enter the operation mode of the app. Calculating the size or hash")]
	[ValidateSet("size", "hash")]
	[String]$mode
)
Function GetFilesInfo {
	param (
		[Parameter(Mandatory = $false)]
        [String]$path,
		[Parameter(Mandatory = $true)]
        [String]$mode,
        [Parameter(Mandatory = $true)]
        [Boolean]$NonInteractive
	)
	try
	{
		if (!$path) {
			# Creating a dialog box form
			$object = New-Object -ComObject Shell.Application
			# Getting the path to the directory being checked
			$path = $object.BrowseForFOlder(0, 'Выберите каталог для которого необходимо сформировать перечень файлов:', 0, 0).Self.Path
		}
		# If the directory is selected and the path exists
		if ($path -and (Test-Path $path))
		{
			# Starting Word
			$Word = New-Object -ComObject Word.Application
			# Word window visible
			$Word.Visible = $false
			# Creating a Word document
			$Document = $Word.Documents.Add()
			# Selection area
			$Selection = $Word.Selection
			# Setting document margins
			$Selection.Pagesetup.TopMargin = 30
			$Selection.Pagesetup.LeftMargin = 30
			$Selection.Pagesetup.RightMargin = 30
			$Selection.Pagesetup.BottomMargin = 30
			# Adding a paragraph
			$Selection.TypeParagraph()
			# Font settings
			$Selection.Font.Name = 'Times New Roman'
			$Selection.Font.Size = 14
			$Selection.ParagraphFormat.Alignment = 1
			$Selection.Font.Bold = $true
			# Adding text to the paragraph
			$Selection.TypeText('Перечень файлов, записанных на CD/DVD диск, инв. № _______')
			# Space before the table
			$Selection.TypeParagraph()
			$Selection.TypeText(' ')
			# Font settings
			$Selection.Font.Name = 'Times New Roman'
			$Selection.Font.Size = 12
			$Selection.ParagraphFormat.Alignment = 3
			$Selection.Font.Bold = $false
			# Adding a table (4 columns)
			$Table = $Word.ActiveDocument.Tables.Add($Word.Selection.Range, 1, 4)
			# Selection border settings in the table
			$Table.Borders.InsideLineStyle = 1
			$Table.Borders.OutsideLineStyle = 1
			# Column widths in the table
			$Table.Columns(1).PreferredWidthType = 2
			$Table.Columns(1).PreferredWidth = 3
			$Table.Columns(3).PreferredWidthType = 2
			$Table.Columns(3).PreferredWidth = 7
			$Table.Columns(4).PreferredWidthType = 2
			$Table.Columns(4).PreferredWidth = 4
			# Table header
			$Table.Cell(1, 1).Range.Text = '№ п/п'
			$Table.Cell(1, 2).Range.Text = 'Название файла/каталога'
			if ($mode -eq "size") {
				$Table.Cell(1, 3).Range.Text = 'Размер файла'		
			}
			else { $Table.Cell(1, 3).Range.Text = 'Hash (MD5)' }
			$Table.Cell(1, 4).Range.Text = 'Гриф'
			# Counters for table generation
			# Rows in the table
			$i = 1
			# Sequential row number
			$counter = 1
			# Flag for directories without files
			$Flag = $false
            # Total sum
            $total = 0
			# List of directories in the selected folder (Sorted by name)
			if (Get-ChildItem -Path $path -Recurse -Directory)
			{
				$dirs = @($path) + (Get-ChildItem -Path $path -Recurse -Directory | Sort-Object Fullname).FullName
			}
			else
			{
				$dirs = @($path)
			}
			# Variable for the progress bar
            $k = 1		
			foreach ($dir in $dirs)
			{
				write-progress -Activity 'Прогресс выполнения:' -PercentComplete ($k/$dirs.Count*100)
				# Adding a new row
				$Table.Rows.Add() | Out-Null
				$i += 1
				# Adding the directory name to the table
				$Table.Cell($i, 2).Range.Text = [string]$dir
				$Table.Cell($i, 2).Range.Font.Bold = $true
				# If there were no files in the previous directory
				if ($Flag)
				{
					# Merging cells in the previous row
					$Table.Cell($i - 1, 2).Merge($Table.Cell($i - 1, 4))
				}
				# Getting the list of files in the directory
				$files = Get-ChildItem -Path $dir -File -Force | Sort-Object Name
				# If there are files in the directory
				if ($files)
				{
					# Changing the flag
					$Flag = $false
					# Marker for the first file
					$j = 0
					foreach ($file in $files)
					{
						# Adding a new row
						$Table.Rows.Add() | Out-Null
						$i += 1
						# If the file is the first in the list
						if ($j -eq 0)
						{
							# Merging cells in the previous row
							$Table.Cell($i - 1, 2).Merge($Table.Cell($i - 1, 4))
						}
						# Adding data to the row
						$Table.Cell($i, 2).Range.Font.Bold = $false
						# Sequential row number
						$Table.Cell($i, 1).Range.Text = [string]$counter
						# File name
						$Table.Cell($i, 2).Range.Text = [string]$file.Name
						if ($mode -eq "size") {
							# File size
							$Table.Cell($i, 3).Range.Text = [string]$file.Length
						}
						else {
							# File hash
							$File_hash = (Get-FileHash -Algorithm MD5 -Path $file.FullName).Hash
							$Table.Cell($i, 3).Range.Text = [string]$file_hash
						}
						# Incrementing counters
						$counter += 1
						$j += 1
                        $total += $file.Length
					}
				}
				# If there are no files in the directory
				else
				{
					# Changing the flag
					$Flag = $true
                    $k += 1
					continue
				}
            $k += 1
			}
            $i += 1
            # Adding a row
			$Table.Rows.Add() | Out-Null
            # Adding the total volume to the row
            $Table.Cell($i, 2).Range.Text = 'Общий объем:'
            $Table.Cell($i, 2).Range.Font.Bold = $true
            $Table.Cell($i, 3).Range.Text = [string]$total
            $Table.Cell($i, 3).Range.Font.Bold = $true
			# Making the font in the table header bold
			$Table.Cell(1, 1).Range.Font.Bold = $true
			$Table.Cell(1, 2).Range.Font.Bold = $true
			$Table.Cell(1, 3).Range.Font.Bold = $true
			$Table.Cell(1, 4).Range.Font.Bold = $true
			# End of table, start a new row
			$Selection.EndKey(6, 0)
			# Path and name of the file to be saved
			$File = "$($PSScriptRoot)\Перечень файлов_"  + (Get-Date -Format 'dd/MM/yyyy HH mm ss') + '.docx'
			# Saving the file
			$Document.SaveAS([ref]$File)		
			if (!$NonInteractive){
				# Displaying a message about successful completion
				[System.Windows.Forms.MessageBox]::Show("Перечень файлов сформирован!", "Готово", [System.Windows.Forms.MessageBoxButtons]::OK)
			}
			else {
				Write-Host $File
			}
		}
		else
		{
			# Displaying a message
			[System.Windows.Forms.MessageBox]::Show('Каталог не выбран или путь не существует!', 'Внимание', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
			return
		}
	}
	# If errors occurred
	catch
	{
		if (!$NonInteractive) {
            # Displaying an error message
            [System.Windows.Forms.MessageBox]::Show("При выполнении программы возникли ошибки: `n $($Error[0])","Внимание",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        }
		return
}
    finally {
        # Closing the document and Word
		$Document.Close()
		$Word.Quit()
    }
return
}
# Determining the script execution mode and performing the function
$NonInteractive = @($False, $True)[$PSBoundParameters.Count -ge 1]
if ($NonInteractive) {
	GetFilesInfo -path $PSBoundParameters.path -mode $PSBoundParameters.mode -NonInteractive $NonInteractive | Out-Null
}
else {GetFilesInfo -mode "size" -NonInteractive $NonInteractive | Out-Null }
