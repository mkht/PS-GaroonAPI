$moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent

Get-Module 'PS-GaroonAPI' | Remove-Module -Force
Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force

Describe 'Tests of Remove-GrnUser' {
    BeforeAll {
        $script:GrnURL = 'https://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
        $ValidCred = New-Object PsCredential 'sato', (ConvertTo-SecureString 'sato' -AsPlainText -Force)
    }

    Context 'Error' {
        It 'If you try to remove user that does not exist, Should -Throw' {
            { Remove-GrnUser -LoginName 'notexistuser' -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Throw '*見つかりません*'
        }
    }

    Context 'Success' {
        BeforeAll {
            $local:Gen_Id = {
                $chars = 'abcdefghkmnprstuvwxyzABCDEFGHKLMNPRSTUVWXYZ123456789'
                $random = 1..8 | ForEach-Object { Get-Random -Maximum $chars.length }
                [string]($chars[$random] -join '')
            }
            $script:UserInfo = @{
                LoginName   = & $Gen_Id
                DisplayName = 'TestUser_X'
                Password    = 'P@ssw0rd'
            }
            New-GrnUser @UserInfo -URL $GrnURL -Credential $ValidCred -ea Stop
        }

        It 'Remove User' {
            { Remove-GrnUser $UserInfo.LoginName -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Not -Throw
        }
    }
}
