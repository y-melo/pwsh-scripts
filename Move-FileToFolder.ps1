<# 
.SYNOPSIS
  Dig through specified path and move each $extension to a folder.
.DESCRIPTION
  Dig through specified path and move each $extension to a folder.
  e.g.:
    Original         > After run the script
    -----------------------------------------
    foo.eml          > foo/foo.eml
    email.eml        > email/email.eml
    Archive/loan.eml > Archive/loan/loan.eml
.NOTES
  File Name  : Move-FileToFolder.ps1
  Author     : Y
  Version    : 1.2.0
  Date       : 2021 Oct 12
  Requires   : Powershel 7.1 or greater
.EXAMPLE
  Move-FileToFolder.ps1 PATH [LIMIT] [WARNING LOG PATH]
  
  Move completed after 0:0:0.
  Total: 1, Moved: 0, Skipped: 1
.EXAMPLE
  Move-FileToFolder.ps1 'D:\Mail Archive'
  
  Move completed after 0:0:0.
  Total: 1, Moved: 0, Skipped: 1
.EXAMPLE
  Move-FileToFolder.ps1 'D:\Mail Archive' '.eml'
  
  Move completed after 0:0:0.
  Total: 1, Moved: 0, Skipped: 1
#>

#Requires -Version 7.1

[CmdletBinding()]
Param (
  [Parameter (Mandatory = $true, Position = 0, HelpMessage = "Parent folder to search for the extension")]
  [string] $parent_folder,
  [Parameter (Mandatory = $false, Position = 1, HelpMessage = "Extension to create folder")]
  [string] $extension = '.eml',
  [Parameter (Mandatory = $false, HelpMessage = "Limit the execution - usseful for testing")]
  [int] $limit = -1,
  [Parameter (Mandatory = $false, HelpMessage = "Path to save warning messages")]
  [string] $warning_log_path = $env:TEMP
)

function Test-ValidFileName {
  param([string]$FileName)

  $IndexOfInvalidChar = $FileName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars())

  # IndexOfAny() returns the value -1 to indicate no such character was found
  return $IndexOfInvalidChar -eq -1
}

function rm_illegal_char {
  Param (
    [Parameter (Mandatory = $true, Position = 0) ]
    [string] $filename 
  )
  Write-Verbose "[rm_illegal_char] FULLNAME: $filename "
  Write-Output "[rm_illegal_char] Removing illegal chars from $filename" >> $warning_log_file
  return $filename -replace '[^\p{L}\p{Nd}/(/)/{/}/_/-/ ]', ''
}

function trim_filename {
  Param (
    [Parameter (Mandatory = $true, Position = 0) ]
    [string] $filename
  )
  
  Write-Verbose "`n[trim_filename] FULLNAME: $filename "
  write-verbose "[trim_filename] $($filename.Trim())"

  return $filename.Trim()

}

function truncate_filename {
  Param(
    [Parameter (Mandatory = $true, Position = 0) ]
    [string] $filename, 
    [int] $filename_max_size = 100
  )
  
  Write-Verbose "[truncate_filename] MAX_SIZE: $filename_max_size"
  Write-Verbose "[truncate_filename] FULLNAME: $filename "
    
  if ($filename.Length -le $filename_max_size) {
    return $filename;
  }
  Write-Debug  "[truncate_filename]: Truncating..."
  Write-Verbose "[truncate_filename] new name: $($filename.Substring(0, $filename_max_size)) "

  return $filename.Substring(0, $filename_max_size)
}

function should_it_move([string]$file) {
  # If parent name is the same as the file return false
  $parent = (Split-Path "$file" -Parent ).Split('\')[-1]
  $flname = trim_filename(truncate_filename(rm_illegal_char([io.path]::GetFileNameWithoutExtension("$file"))))
  $total_length = ($file.Length + $flname.Length + 4)
  Write-Verbose "[should_it_move] PARENT: $parent"
  Write-Verbose "[should_it_move] FLNAME: $flname"
  Write-Verbose "[should_it_move] Total Length: $total_length ($($file.Length)+$($parent.Length))"

  if ( $parent.Equals(($flname.Trim())) ) {
    Write-Verbose "[should_it_move] RETURN: FALSE`n`n"
    return $false
  }

  
  write-verbose "[should_it_move] RETURN: TRUE`n`n"
  if ($total_length -ge 255) {
    $mod_length = ($total_length - 255)
    Write-Debug "[should_it_move]If moved it will be above the limit of 255 char. Need to remove $mod_length char(s)"
    Write-Output "Need to remove $mod_length chars from '$file'" >> $warning_log_file
  }
  
  return $true;
}

function create_folder([string]$file) {
  $parent = Split-Path "$file" -Parent
  $folder = ([io.path]::GetFileNameWithoutExtension("$file")).Trim()
  Write-Verbose "[create_folder] Length: $($folder.Length). Folder Lenght too big will be truncated"
  $folder = trim_filename(truncate_filename(rm_illegal_char($folder)))
  $file_fullpath = "$parent\$folder"
  Write-Verbose "[create_folder] PARENT: $PARENT"
  Write-Verbose "[create_folder] FOLDER: $folder"
  Write-Verbose "[create_folder] Moving $file"
  Write-Verbose "[create_folder] To: $file_fullpath"

  New-Item -ItemType Directory -Path "$file_fullpath" -Force | Out-Null
    
  Write-Debug "[create_folder] Move-Item '$file' '$file_fullpath' -Force"
  Move-Item -LiteralPath "$file" -Destination "$file_fullpath" -Force 

}

## CODE ##
#
#

# Vars
$counter = 0    # Increments for each file
$skipped_counter = 0    # Increments for each skipped file
$moved_counter = 0    # Increment for each moved file
$warning_log_file = "$warning_log_path\warning.log"
##

Write-Output "" > $warning_log_file

$files = (Get-ChildItem -Path $parent_folder -Filter "*$extension" -Recurse -ErrorAction SilentlyContinue -Force | ForEach-Object { $_.FullName })

Write-Output "FOLDER TO SEARCH: $parent_folder | Extension: $extension"
Write-Output "Found: $($files.Count) '$extension' files."
Start-Sleep 2

[timespan] $executionTime = Measure-Command {
  foreach ($file in $files) {
    Write-Progress -Activity "Creating Folder $file" -Status "$($counter) of $($files.Count)" -PercentComplete ($counter / $($files.Count) * 100)
  
    Write-Debug "[main] Item: $file"
    
    if (should_it_move("$file")) {
      create_folder("$file")
      Write-Debug "[main] COPYING $file"
      $moved_counter++
    }
    else {
      Write-Debug "[main] SKIPPING $file"
      $skipped_counter++
    }
    $counter++
    if ($limit -gt 0 -and $counter -ge $limit ) {
      Write-Debug "[main] Breaking... limit of $limit reached ($counter)"
      break;
    }
  }
}

$elapsed = "$($executionTime.Hours):$($executionTime.Minutes):$($executionTime.Seconds)"
Write-Output "Move completed after $elapsed.`nTotal: $counter, Moved: $moved_counter, Skipped: $skipped_counter"
if ($skipped_counter -gt 0) {
  Write-Output "SKipped files saved: $warning_log_file"
}

exit 0
