#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.3' }

BeforeAll {
    $moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Get-Module 'PS-GaroonAPI' | Remove-Module -Force
    Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force
}

AfterAll {
    Get-Module 'PS-GaroonAPI' | Remove-Module -Force
}

Describe 'Tests of Get-GrnUser' {
    BeforeAll {
        $GrnURL = 'https://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
        $ValidCred = New-Object PsCredential 'sato', (ConvertTo-SecureString 'sato' -AsPlainText -Force)
    }

    Context 'ユーザを取得する' {
        BeforeAll {
            $script:User1 = $null
        }

        It '実行時にエラーが発生しないか' {
            { $script:User1 = Get-GrnUser -LoginName 'nomura' -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Not -Throw
        }

        It '正しいユーザが取得できているか' {
            $User1.DisplayName | Should -Be '野村 理佳'
            $User1.Kana | Should -Be 'のむら　りか'
            $User1.Email | Should -Be 'nomura@localhost'
            $User1.Position | Should -Be 10
            @($User1.Organization).Count | Should -Be 3
            $User1.PrimaryOrganization | Should -Be '営業部'
            $User1.Invalid | Should -Be $false
        }

        AfterAll {
            $script:User1 = $null
        }
    }

    Context '複数のユーザ情報を取得する（パイプライン）' {
        BeforeAll {
            $script:User1 = $null
        }

        It '実行時にエラーが発生しないか' {
            { $script:Users = ('yamada', 'John') | Get-GrnUser -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Not -Throw
        }

        It '複数ユーザが取得できているか' {
            $Users -is [Array] | Should -Be $true
            @($Users).Count | Should -Be 2
        }

        It '正しいユーザが取得できているか' {
            $Users[0].DisplayName | Should -Be '山田 隆志'
            $Users[1].DisplayName | Should -Be 'John Doe'
        }

        AfterAll {
            $script:User1 = $null
        }
    }

    Context '存在しないユーザは取得できない' {
        BeforeAll {
            $script:User2 = 'This variable must be null'
        }

        It '実行時にエラーが発生しないか' {
            { $script:User2 = Get-GrnUser 'nouser' -URL $GrnURL -Credential $ValidCred -ea Stop -wa SilentlyContinue } | Should -Not -Throw
        }

        It '関数が明示的に$nullを返しているか' {
            $User2 | Should -BeNullOrEmpty
            @($User2).Count | Should -Be 1   #明示的な$nullの場合は1になるが、何も値を持っていない暗黙$nullの場合は0になる
        }

        AfterAll {
            $script:User2 = $null
        }
    }
}
