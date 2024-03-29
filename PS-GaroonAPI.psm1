﻿# Load Classes
$ClassList = @(
    'GaroonClass.ps1',
    'GaroonBase.ps1',
    'GaroonAdmin.ps1',
    'GaroonMail.ps1',
    'GaroonAddress.ps1'
)
foreach ($class in $ClassList) {
    . $PSScriptRoot\Class\$class
}

# Load Functions
$AllFunctions = @( Get-ChildItem -Path $PSScriptRoot\Function\ -Filter *.ps1 -Recurse -File -ErrorAction SilentlyContinue )
foreach ($function in $AllFunctions) {
    try {
        # Dot source
        . $function.fullname
    }
    catch {
        Write-Error -Message "Failed to import function $($function.fullname): $_"
    }
}

Export-ModuleMember -Function $AllFunctions.Basename