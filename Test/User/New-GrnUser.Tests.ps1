﻿#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.3' }

BeforeAll {
    $moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Get-Module 'PS-GaroonAPI' | Remove-Module -Force
    Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force
}

AfterAll {
    Get-Module 'PS-GaroonAPI' | Remove-Module -Force
}

Describe 'Tests of New-GrnUser' {
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
            { New-GrnUser @DummyInfo -URL $WrongURL -Credential $ValidCred -ea Stop } | Should -Throw '*404*'
        }

        It 'If you use invalid credential, Should -Throw [FW00007] error' {
            $local:InvalidCred = New-Object PsCredential 'hoge', (ConvertTo-SecureString 'hoge' -AsPlainText -Force)
            { New-GrnUser @DummyInfo -URL $GrnURL -Credential $InvalidCred -ea Stop } | Should -Throw '*[FW00007]*'
        }

        It 'If try to create duplicate user, Should -Throw' {
            $local:DuplicateUser = @{
                LoginName   = 'sato'
                Password    = 'sato'
                DisplayName = '佐藤 昇'
            }
            { New-GrnUser @DuplicateUser -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Throw '*既に存在します*'
        }

        It 'If specified NON existing Organization name, Should -Throw' {
            $local:DuplicateUser = $DummyInfo
            $DuplicateUser.PrimaryOrganization = 'ほにゃららら'
            $DuplicateUser.Organization = '役員'
            { New-GrnUser @DuplicateUser -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Throw '*存在しない組織名が含まれています*'
        }
    }

    Context 'Success' {
        BeforeAll {
            $script:Gen_Id = {
                $chars = 'abcdefghkmnprstuvwxyzABCDEFGHKLMNPRSTUVWXYZ123456789'
                $random = 1..8 | ForEach-Object { Get-Random -Maximum $chars.length }
                [string]($chars[$random] -join '')
            }
        }

        It 'Create new user with orgs' {
            $script:UserInfo = @{
                LoginName           = & $Gen_Id
                DisplayName         = '手巣戸 太郎'
                Password            = 'P@ssw0rd'
                Kana                = 'てすと たろう'
                PrimaryOrganization = '役員'
                Organization        = '総務部', '経理部'
            }

            $private:user = New-GrnUser @UserInfo -URL $GrnURL -Credential $ValidCred -PassThru -ea Continue
            $user.LoginName | Should -BeExactly $UserInfo.LoginName
            $user.DisplayName | Should -BeExactly $UserInfo.DisplayName
            $user.PrimaryOrganization | Should -BeExactly $UserInfo.PrimaryOrganization
            $user.Organization -ccontains $UserInfo.PrimaryOrganization | Should -Be $true
            $UserInfo.Organization.ForEach({
                    $user.Organization -ccontains $_ | Should -Be $true
                })
        }

        It 'Create new user without orgs' {
            $script:UserInfo = @{
                LoginName           = & $Gen_Id
                DisplayName         = '手巣戸 花子'
                Password            = 'P@ssw0rd'
                PrimaryOrganization = ''
                Organization        = ''
            }

            $private:user = New-GrnUser @UserInfo -URL $GrnURL -Credential $ValidCred -PassThru -ea Continue
            $user.LoginName | Should -BeExactly $UserInfo.LoginName
            $user.DisplayName | Should -BeExactly $UserInfo.DisplayName
            $UserInfo.Organization | Should -Be ''
        }

        It 'Create new user contains characters that need to be escaped' {
            $script:UserInfo = @{
                LoginName           = & $Gen_Id
                DisplayName         = 'Smith "<Agent>" Jones & Rude'
                Password            = 'P@ssw0rd&X'
                PrimaryOrganization = ''
                Organization        = ''
            }

            $private:user = New-GrnUser @UserInfo -URL $GrnURL -Credential $ValidCred -PassThru -ea Continue
            $user.LoginName | Should -BeExactly $UserInfo.LoginName
            $user.DisplayName | Should -BeExactly $UserInfo.DisplayName
            $UserInfo.Organization | Should -Be ''
        }

        AfterEach {
            if ($UserInfo) {
                Set-GrnUser -LoginName $UserInfo.LoginName -Invalid $true -URL $GrnURL -Credential $ValidCred -ea Continue
            }
        }
    }
}
