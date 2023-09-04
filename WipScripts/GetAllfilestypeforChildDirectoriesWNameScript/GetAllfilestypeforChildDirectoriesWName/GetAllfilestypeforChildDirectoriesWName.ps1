# Specify the drive on which you want to perform the search
$drive = "D:\"

# Search for child directories with the name "wip"
$wipDirectories = Get-ChildItem -Path $drive -Recurse | Where-Object { $_.PSIsContainer -and $_.Name -eq "wip" -and $_.FullName -notlike 'C:\Windows*' }
# Traverse through the "wip" directories and check for XLS files
foreach ($directory in $wipDirectories) {
    $directoryPath = $directory.FullName
    $xlsFiles = Get-ChildItem -Path $directoryPath -Filter "*.xls" -File

    if ($xlsFiles.Count -gt 0) {
        Write-Host "Directory '$directoryPath' contains XLS files:"
        $xlsFiles | ForEach-Object {
            Write-Host " - $($_.Name)"
        }
    } else {
        Write-Host "Directory '$directoryPath' does not contain any XLS files."
    }
}