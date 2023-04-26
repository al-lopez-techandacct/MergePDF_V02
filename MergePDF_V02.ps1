<#
This module provides functionality to merge multiple pdf files. It was tested with PdfSharp-gdi.dll in version 1.5.4 and powershell 5 and is provided "as is"
#>
function MergePdf {
<#
.SYNOPSIS
Merges multiple PDF files into a single multisheet PDF.
.PARAMETER Files
A collection of PDF files
.PARAMETER DestinationFile
Full path of the merged PDF
.PARAMETER PdfSharpPath
Full path to the required PdfSharp library
.PARAMETER Force
Tries to force destination file and directory creation and deletion of source files, even when they are read-only
.PARAMETER
RemoveSourceFiles
Deletes the source files after PDF is merged
.EXAMPLE
$files = Get-ChildItem "C:\temp\PDF\Source" -Filter "*.pdf"
MergePdf -Files $files  -DestinationFile "C:\TEMP\PDF\Destination\test.pdf" -PdfSharpPath 'C:\ProgramData\coolOrange\powerJobs\Modules\PdfSharp-gdi.dll' -Force -RemoveSourceFiles
#>
param(
[Parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]
[ValidateScript({
    if( $_.Extension -ine ".pdf" ){ 
        throw "The file $($_.FullName) is not a pdf file."
    } 
    if(-not (Test-Path $_.FullName)) {
        throw "The file '$($_.FullName)' does not exist!"
    }
    $true
})]
[System.IO.FileInfo[]]$Files,
[Parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]
[System.IO.FileInfo]$DestinationFile,
[Parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]
$PdfSharpPath,
[switch]$Force,
[switch]$RemoveSourceFiles
)
    Write-Host ">> $($MyInvocation.MyCommand.Name) >>"

    if((Test-Path $PdfSharpPath) -eq $false) {
        throw "Could not find pdfsharp assembly at $($PdfSharpPath)"
    }
    Add-Type -LiteralPath $PdfSharpPath

    if((Test-Path $DestinationFile.FullName) -and $DestinationFile.IsReadOnly -and -not $Force) {
        throw "Destination file '$($DestinationFile.FullName)' is read only"
    }

    [System.IO.DirectoryInfo]$DestinationDirectory = $DestinationFile | Split-Path -Parent
    if(-not (Test-Path $DestinationDirectory)) {
        try {
            $DestinationDirectory = New-Item -Path $DestinationDirectory.FullName -ItemType Directory -Force:$Force
        } catch {
            throw "Error in $($MyInvocation.MyCommand.Name). Could not create directory '$($Path)'. $Error[0]"
        }
    }

    $pdf = New-Object PdfSharp.Pdf.PdfDocument
    Write-Host "Creating new PDF"
    foreach ($file in $Files) {
        $inputDocument = [PdfSharp.Pdf.IO.PdfReader]::Open($file.FullName, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import)
        for ($index = 0; $index -lt $inputDocument.PageCount; $index++) {
            $page = $inputDocument.Pages[$index]
            $null = $pdf.AddPage($page)
        }
    }

    Write-Host "Saving PDF"
    if((Test-Path $DestinationFile.FullName) -and $Force) { 
        Remove-Item $DestinationFile.FullName -Force 
    }
    $pdf.Save($DestinationFile.FullName)

    if($RemoveSourceFiles) {
        Write-Host "Removing source files"
        foreach($file in $files) {
            Remove-Item -Path $file.FullName -Force:$Force
        }
    }
}


function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton)
{
    $browseForFolderOptions = 0
    if ($NoNewFolderButton) { $browseForFolderOptions += 512 }

    $app = New-Object -ComObject Shell.Application
    $folder = $app.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
    if ($folder) { $selectedDirectory = $folder.Self.Path } else { $selectedDirectory = '' }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) > $null
    return $selectedDirectory
}


#####  Script processing starts here #####

$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
#Point to the folder that holds the .pdf files to be combined.
$directoryPath = Read-FolderBrowserDialog -Message "Please select a directory" -InitialDirectory 'c:\' -NoNewFolderButton
if (![string]::IsNullOrEmpty($directoryPath))
    { Write-Host "You selected the directory: $directoryPath" }
else 
    { "You did not select a directory." }

#Prompt for the name of the combined .pdf file.
Add-Type -AssemblyName Microsoft.VisualBasic
$NewPdfFilename = [Microsoft.VisualBasic.Interaction]::InputBox('Name of the combined file (Do not add.pdf extension')
$CreatePdfFile = "$directoryPath\$NewPdfFilename.pdf"

#$files = Get-ChildItem "C:\Downloads\22-May-2021.1317222.invoice.BNCInsurance" -Filter "*.pdf"
#MergePdf -Files $files  -DestinationFile "C:\Downloads\22-May-2021.1317222.invoice.BNCInsurance\22-May-2021.1317222.invoice.BNCInsurance.NOD.Detail.pdf" -PdfSharpPath '.\PdfSharp-gdi.dll' -Force

$files = Get-ChildItem $directoryPath -Filter "*.pdf"
$SharpDllPath = $ScriptDir + '\PdfSharp-gdi.dll'
MergePdf -Files $files -DestinationFile "$CreatePdfFile" -PdfSharpPath $SharpDllPath -Force