# Author: Giancarlo Trevisan
# Date: 2025-06-05
# Description: 
#   This script increments the patch version in FillODT.csproj, publishes the project for win-x64,
#   renames and zips the output folder, and moves the zip to the current directory.

# 1. Increment patch version in FillODT.csproj
$csproj = Join-Path $pwd.Path "FillODT.csproj"
[xml]$proj = Get-Content $csproj -Raw

# Find the Version node
$versionNode = $proj.GetElementsByTagName("Version")[0]
$version = $versionNode.InnerText

$parts = $version.split('.')
$parts[2] = ([int]$parts[2] + 1).ToString()
$newVersion = "$($parts[0]).$($parts[1]).$($parts[2])"
$versionNode.InnerText = $newVersion
$proj.Save($csproj)

# 2. Publish
dotnet publish -c Release -r win-x64

# 3. Rename publish folder
$publishDir = Join-Path $pwd.Path "\bin\Release\net9.0\win-x64\publish"
$targetDir = "FillODT.v$newVersion.win-x64"
if (Test-Path $publishDir.Replace("publish", $targetDir)) { 
	Remove-Item $publishDir.Replace("publish", $targetDir) -Recurse -Force 
}
Rename-Item $publishDir $targetDir
$publishDir = $publishDir.Replace("publish", $targetDir)

# 4. Zip the folder
$zipFile = "$publishDir.zip"
if (Test-Path $zipFile) { 
	Remove-Item $zipFile -Force 
}
Compress-Archive -Path $publishDir\* -DestinationPath $zipFile

# 5. Clean up
Remove-Item $publishDir -Recurse -Force
Move-Item -Path $zipFile -Destination $pwd
