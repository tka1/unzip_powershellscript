PARAM (
    [string] $ZipFilesPath = "c:\temp\rbndata",
    [string] $UnzipPath = "c:\temp\rbndata\csv"
)
 
$Shell = New-Object -com Shell.Application
$Location = $Shell.NameSpace($UnzipPath)
 
$ZipFiles = Get-Childitem $ZipFilesPath -Recurse -Include *.ZIP
 
$progress = 1
foreach ($ZipFile in $ZipFiles) {
    Write-Progress -Activity "Unzipping to $($UnzipPath)" -PercentComplete (($progress / ($ZipFiles.Count + 1)) * 100) -CurrentOperation $ZipFile.FullName -Status "File $($Progress) of $($ZipFiles.Count)"
    $ZipFolder = $Shell.NameSpace($ZipFile.fullname)
 
 
    $Location.Copyhere($ZipFolder.items(), 1040) # 1040 - No msgboxes to the user - http://msdn.microsoft.com/en-us/library/bb787866%28VS.85%29.aspx
    $progress++
}