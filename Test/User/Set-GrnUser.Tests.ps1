#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.3' }

BeforeAll {
    $moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Get-Module 'PS-GaroonAPI' | Remove-Module -Force
    Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force
}

AfterAll {
    Get-Module 'PS-GaroonAPI' | Remove-Module -Force
}

Describe 'Tests of Set-GrnUser' {
    BeforeAll {
        $script:GrnURL = 'https://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
        $ValidCred = New-Object PsCredential 'sato', (ConvertTo-SecureString 'sato' -AsPlainText -Force)
    }

    Context 'Error' {
        BeforeAll {
            $script:DummyInfo = @{
                LoginName   = 'foo'
                Password    = 'pass'
                DisplayName = 'bar'
            }
        }

        It 'If http 404 error, Should -Throw WebException' {
            $local:WrongURL = 'https://onlinedemo2.cybozu.info/scripts/garooooooooooon/grn.exe'
            { Set-GrnUser @DummyInfo -URL $WrongURL -Credential $ValidCred -ea Stop } | Should -Throw '*404*'
        }

        It 'If you use invalid credential, Should -Throw [FW00007] error' {
            $local:InvalidCred = New-Object PsCredential 'hoge', (ConvertTo-SecureString 'hoge' -AsPlainText -Force)
            { Set-GrnUser @DummyInfo -URL $GrnURL -Credential $InvalidCred -ea Stop } | Should -Throw '*[FW00007]*'
        }

        It 'If specified NON existing user, Should -Throw' {
            $local:NonExistingUser = @{
                LoginName = 'hogehoge'
            }
            { Set-GrnUser @NonExistingUser -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Throw '*[GRN_SYSAPI_64008]*'
        }

        It 'If specified NON existing Organization name, Should -Throw' {
            $local:NonExistingOrgs = @{
                LoginName = 'nomura'
            }
            $NonExistingOrgs.PrimaryOrganization = 'ほにゃららら'
            $NonExistingOrgs.Organization = '役員'

            { Set-GrnUser @NonExistingOrgs -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Throw '*存在しない組織名が含まれています*'
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
                LoginName           = & $Gen_Id
                DisplayName         = 'TestUser_X'
                Password            = 'P@ssw0rd'
                PrimaryOrganization = '役員'
                Organization        = '役員', '経理部'
            }
            New-GrnUser @UserInfo -URL $GrnURL -Credential $ValidCred -ea Continue
        }

        AfterAll {
            if ($UserInfo) {
                Set-GrnUser -LoginName $UserInfo.LoginName -Invalid $true -URL $GrnURL -Credential $ValidCred -ea Continue
            }
        }

        It 'Modify user properties (DisplayName,Phone)' {
            $private:info = @{
                DisplayName = 'Modify_X'; Phone = '080-XXX-XXXX'
            }
            $private:user = Set-GrnUser -LoginName $UserInfo.LoginName @info -URL $GrnURL -Credential $ValidCred -PassThru -ea Continue
            $user.LoginName | Should -BeExactly $UserInfo.LoginName
            $user.DisplayName | Should -BeExactly $info.DisplayName
            $user.Phone | Should -BeExactly $info.Phone
        }

        It 'Modify user Orgs' {
            $private:info = @{
                PrimaryOrganization = '社長'; Organization = '社長', '企画部'
            }
            $private:user = Set-GrnUser -LoginName $UserInfo.LoginName @info -URL $GrnURL -Credential $ValidCred -PassThru -ea Continue
            $user.LoginName | Should -BeExactly $UserInfo.LoginName
            $user.PrimaryOrganization | Should -BeExactly $info.PrimaryOrganization
            $user.Organization -ccontains $info.PrimaryOrganization | Should -Be $true
            $info.Organization.ForEach({
                    $user.Organization -ccontains $_  | Should -Be $true
                })
        }

        It 'Clear user Orgs' {
            $private:info = @{
                Organization = ''
            }
            $private:user = Set-GrnUser -LoginName $UserInfo.LoginName @info -URL $GrnURL -Credential $ValidCred -PassThru -ea Continue -WarningAction SilentlyContinue
            $user.LoginName | Should -BeExactly $UserInfo.LoginName
            $info.Organization | Should -Be ''
        }

        It 'Pipeline from Get-GrnUser' {
            $private:info = @{
                Organization        = '総務部', '開発部'
                PrimaryOrganization = '情報システム部'
                Position            = 30347
            }
            $private:user = Get-GrnUser -LoginName $UserInfo.LoginName -URL $GrnURL -Credential $ValidCred |
            Set-GrnUser @info -URL $GrnURL -Credential $ValidCred -PassThru -ea Continue
            $user.LoginName | Should -BeExactly $UserInfo.LoginName
            $user.Position | Should -BeExactly $info.Position
            $user.PrimaryOrganization | Should -BeExactly $info.PrimaryOrganization
            $user.Organization -ccontains $info.PrimaryOrganization | Should -Be $true
            $info.Organization.ForEach({
                    $user.Organization -ccontains $_ | Should -Be $true
                })
        }

    }
}
