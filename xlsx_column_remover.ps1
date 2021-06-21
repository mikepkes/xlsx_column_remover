# Init PowerShell Gui
Add-Type -AssemblyName System.Windows.Forms
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')


Write-Host "Requesting Root Source Folder From User..."

$RootDirectoryBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$RootDirectoryBrowser.Description = "Select Root Source Folder to Work On"
$null = $RootDirectoryBrowser.ShowDialog()


Write-Host "User selected path for root work:" $RootDirectoryBrowser.SelectedPath

Write-Host "Requesting Output Folder From User..."

$OutputDirectoryBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$OutputDirectoryBrowser.Description = "Select Output Folder"
$null = $OutputDirectoryBrowser.ShowDialog()

Write-Host "User selected path for output:" $OutputDirectoryBrowser.SelectedPath

$title = 'Column Letters'
$msg   = 'Choose Columns to Remove separated by a comma.  IE: "D,E,G"'

$Columns_Input = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
Write-Host $Columns_Input
$Columns_Input.toUpper()
$ColumnsToRemove = $Columns_Input -split ','

$ColumnsToRemove = $ColumnsToRemove | Sort-Object -Descending

Write-Host "Columns to Remove from Copied XLSX Files:" $ColumnsToRemove

Get-ChildItem -Path $RootDirectoryBrowser.SelectedPath -Filter *.xlsx -Recurse -File -Name| ForEach-Object {
    $rel_folder = Split-Path -Path $_
    $xlsx_file = Join-Path -Path $RootDirectoryBrowser.SelectedPath -ChildPath  $_

    $target_folder = Join-Path -Path $OutputDirectoryBrowser.SelectedPath -ChildPath $rel_folder
    $target_xlsx = Join-Path -Path $OutputDirectoryBrowser.SelectedPath -ChildPath $_

    if (Test-Path -Path $target_folder) {
        Write-Host $target_folder "already exists, not creating it."
    }
    else {
        Write-Host "Creating output folder:" $target_folder
        New-Item -Path $target_folder -ItemType Directory
    }

    Write-Host "Copying Original XLSX file to Output Path"
    Copy-Item -Path $xlsx_file -Destination $target_folder

    Write-Host "Opening Instance of Excel"
    $excel = New-Object -ComObject Excel.Application -Property @{Visible = $false} 

    Write-Host "Using Excel to open file:" $target_xlsx

    $excel.Workbooks.Open($target_xlsx) | ForEach-Object -Process {
        $wb = $_
        $wb.Worksheets | ForEach-Object -Process {
            # Delete Column
            # $_.Activate
            Write-Host "Acting on Worksheet:" $_.Name

            Write-Host "All Column Removals" $ColumnsToRemove
            foreach ($column in $ColumnsToRemove) {
                Write-Host "Removing Column:" $column
                $current_column_range = "${column}:${column}"
                $_.Range($current_column_range).EntireColumn.Delete()
            }

        }

        Write-Host "Saving File:" $target_xlsx
        $wb.Save()
    }
    

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    Remove-Variable excel



}



Write-Host -NoNewLine 'Process Complete: Press any key to continue...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

