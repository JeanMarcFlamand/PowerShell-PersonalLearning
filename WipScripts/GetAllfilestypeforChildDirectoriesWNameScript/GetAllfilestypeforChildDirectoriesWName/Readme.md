# PowerShell Directory Search Script

This PowerShell script is designed to search for child directories with the name "wip" within a specified drive and its subfolders. It also checks if these "wip" directories contain any files with the ".xlsx" extension. Here is a summary of what the script does:

## Usage

1. Specify the drive on which you want to perform the search by setting the `$drive` variable to your desired drive letter or path.

2. The script searches for child directories named "wip" within 4 levels of subfolders from the specified drive, excluding system directories like 'C:\Windows*'.

3. It checks if any "wip" directories are found within the search criteria.

4. If "wip" directories are found, the script traverses through each of them and checks for XLS files.

5. If XLS files are present within a "wip" directory, it lists them.

6. The script also includes a progress bar to track the search progress.

## Running the Script

Ensure you have PowerShell installed on your system. Open a PowerShell terminal, paste the script, and run it. Adjust the `$drive` variable to specify your desired drive or path for the search.

**Note:** This script provides a convenient way to search for specific directories and files within a drive's subfolders. Make sure to review and customize it according to your requirements.
