Write-Host 'Build Start'

$Src_Class = Join-Path $PSScriptRoot '..\Class\'
$Src_Func = Join-Path $PSScriptRoot '..\Function\'
$Dest = Join-Path $PSScriptRoot '..\'

Get-Content @(
    (Join-Path $Src_Class ".\GaroonClass.ps1"),
    (Join-Path $Src_Class ".\GaroonMail.ps1"),
    (Join-Path $Src_Class ".\GaroonBase.ps1"),
    (Join-Path $Src_Class ".\GaroonAdmin.ps1"),
    (Join-Path $Src_Func ".\Get-GrnOrganization.ps1"),
    (Join-Path $Src_Func ".\New-GrnOrganization.ps1"),
    (Join-Path $Src_Func ".\Get-GrnUser.ps1"),
    (Join-Path $Src_Func ".\New-GrnUser.ps1"),
    (Join-Path $Src_Func ".\Set-GrnUser.ps1"),
    (Join-Path $Src_Func ".\Remove-GrnUser.ps1"),
    (Join-Path $Src_Func ".\Invoke-SOAPRequest.ps1")
) |
    ? {($_ -notmatch '^\s*#+(?!>)')} |
    ? {($_ -notmatch '^\s*$')} |
    ? {$_ -notmatch '^\s*using'} |
    Set-Content (Join-Path $Dest 'PS-GaroonAPI.psm1') -Force -Encoding UTF8

Write-Host 'Build Complete'