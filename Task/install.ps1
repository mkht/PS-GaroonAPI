$Src = Join-Path $PSScriptRoot '..\'
$Dst = "$env:SystemDrive\Program Files\WindowsPowerShell\Modules"
$ModuleName = "PS-GaroonAPI"

if(Get-ChildItem $Src -Filter "*.psm1"){
    if(Test-Path($Dst)){
        New-Item -Path (Join-Path $Dst $ModuleName) -Value $Src -ItemType SymbolicLink -Force -Verbose
        Write-Host 'Success'
        return
    }
}

Write-Warning "Fail"