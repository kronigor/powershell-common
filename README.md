# PowerShell Scripts Repository
This repository contains a collection of PowerShell scripts designed to perform various tasks related to system administration, information retrieval from Active Directory (AD), and data processing. Each script resides in its own directory, named after the script for easy navigation and use. Below is a brief overview of each script along with examples of how to run them.

## Scripts

### 1. Get Computers from AD (`get-computers-from-ad.ps1`)

Fetches a list of computer objects from Active Directory.

#### Parameters
- `organizational_units`: Enter an Organizational Unit. Allows multiple entries separated by commas. Example: "OU=Computers,DC=example,DC=com".
- `build_number`: Enter an Operating System build number to filter computers. Example: "1909".

#### Example
```
powershell.exe .\get-computers-from-ad.ps1 -ou "OU=Computers,DC=example,DC=com" -bn "1909"
```

### 2. Get Smartcard Logon from AD (`get-smartcardlogon-from-ad.ps1`)

Queries Active Directory for users required to use a smartcard for logon.

#### Parameters
- `organizational_units`: Specifies the Organizational Unit(s) to search within. Example: "OU=Users,DC=example,DC=com".

#### Example
```
powershell.exe .\get-smartcardlogon-from-ad.ps1 -ou "OU=Users,DC=example,DC=com"
```

### 3. Get Files Info from Directory (`get-filesinfo-from-directory.ps1`)

Generates a detailed report of files in a specified directory.

#### Parameters
- `path`: The path of the directory to scan for files.
- `mode`: The operation mode of the script. Options are "size" for calculating file size or "hash" for generating file hashes.

#### Example
```
powershell.exe .\get-filesinfo-from-directory.ps1 -p "C:\Your\Directory\Path" -m "hash"
```

### 4. Get Software from ARMS to XLSX (`get-soft-from-arms-to-xlsx.ps1`)

Exports a list of software from the Asset Management System (ARMS) to an Excel file.

#### Parameters
- `OUs`: Organizational Unit(s) for the software list extraction. Example: "OU=Workstations,DC=example,DC=com".
- `dir`: Directory for temporary JSON file storage during processing.
- `maxThreads`: The maximum number of concurrent operations for software list extraction.
- `packetCount`: Number of ping requests to send for each workstation connectivity test.
- `armNames`: Adds the workstation names to the report.

#### Example
```
powershell.exe .\get-soft-from-arms-to-xlsx.ps1 -ou "OU=Workstations,DC=example,DC=com" -d "C:\Temp\ARMS" -t 15 -c 2 -a $True
```

## Getting Started

To run these scripts, you'll need PowerShell installed on your computer. Clone this repository to your local machine, navigate to the script's directory, and run the script using PowerShell. Ensure you have the necessary permissions, especially for scripts interacting with Active Directory and the file system.
```shell
git clone https://github.com/kronigor/powershell-common.git
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.
