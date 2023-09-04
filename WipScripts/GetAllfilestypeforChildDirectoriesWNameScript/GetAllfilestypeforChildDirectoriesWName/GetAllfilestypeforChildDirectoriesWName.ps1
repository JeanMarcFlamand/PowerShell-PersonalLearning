# Specify the drive on which you want to perform the search
$drive = "D:\"

# Search for child directories with the name "wip"
$wipDirectories = Get-ChildItem -Path $drive -Recurse -Directory -Depth 4 | Where-Object { $_.PSIsContainer -and $_.Name -eq "wip" -and $_.FullName -notlike 'C:\Windows*' }
# Check if $wipDirectories is empty or null
if ($wipDirectories -eq $null -or $wipDirectories.Count -eq 0)
{
    Write-Host "No 'wip' directories found within 4 levels of subfolders from $drive."
}
else
{
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
    # Increment the progress counter and update the progress bar
        $progressCounter++
        Write-Progress -PercentComplete (($progressCounter / $totalWipDirectories) * 100) -Status "Progress" -CurrentOperation "Processing" -Id 1
    }

    # Complete the progress bar
    Write-Progress -Completed -Status "Completed" -Id 1
}
# Prompt the user to press any key to continue and prevent window closure
Write-Host "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")