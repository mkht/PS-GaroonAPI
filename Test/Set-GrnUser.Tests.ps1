﻿$moduleRoot = Split-Path $PSScriptRoot -Parent

Get-Module 'PS-GaroonAPI' | Remove-Module -Force
Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force

Describe "Tests of Set-GrnUser" {
    $script:GrnURL = 'http://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
    $ValidCred = New-Object PsCredential "sato",(ConvertTo-SecureString "sato" -asplaintext -force)
    
    Context 'Error' {
        
        $script:DummyInfo = @{
                LoginName = 'foo'
                Password = 'pass'
                DisplayName = 'bar'
            }

        It "If http 404 error, shoud throw WebException" {
            $local:WrongURL = 'http://onlinedemo2.cybozu.info/scripts/garooooooooooon/grn.exe'

            {Set-GrnUser @DummyInfo -URL $WrongURL -Credential $ValidCred -ea Stop} | Should throw '404'
        }

        It "If you use invalid credential, shoud throw [FW00007] error" {
            $local:InvalidCred = New-Object PsCredential "hoge",(ConvertTo-SecureString "hoge" -asplaintext -force)
            
            {Set-GrnUser @DummyInfo -URL $GrnURL -Credential $InvalidCred -ea Stop} | Should throw '[FW00007]'
        }

        It "If specified NON existing user, shoud throw" {
            $local:NonExistingUser = @{
                LoginName = 'hogehoge'
            }
            {Set-GrnUser @NonExistingUser -URL $GrnURL -Credential $ValidCred -ea Stop} | Should throw 'ユーザーが見つかりません'
        }

        It "If specified NON existing Organization name, shoud throw" {
            
            $local:NonExistingOrgs = @{
                LoginName = 'nomura'
            }
            $NonExistingOrgs.PrimaryOrganization = 'ほにゃららら'
            $NonExistingOrgs.Organization = '役員'
            
            {Set-GrnUser @NonExistingOrgs -URL $GrnURL -Credential $ValidCred -ea Stop} | Should throw '存在しない組織名が含まれています'
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
                LoginName = & $Gen_Id
                DisplayName = 'TestUser_X'
                Password = 'P@ssw0rd'
                PrimaryOrganization = '役員'
                Organization = "役員","経理部"
            }
            New-GrnUser @UserInfo -URL $GrnURL -Credential $ValidCred -ea Continue
        }

        AfterAll {
            if($UserInfo){
                Set-GrnUser -LoginName $UserInfo.LoginName -Invalid $true -URL $GrnURL -Credential $ValidCred -ea Continue
            }
        }

        It "Modify user properties (DisplayName,Phone)" {
            $private:info = @{
                DisplayName = 'Modify_X'; Phone = '080-XXX-XXXX'
            }
            $private:user = Set-GrnUser -LoginName $UserInfo.LoginName @info -URL $GrnURL -Credential $ValidCred -PassThru -ea Continue            
            $user.LoginName | Should BeExactly $UserInfo.LoginName
            $user.DisplayName | Should BeExactly $info.DisplayName
            $user.Phone | Should BeExactly $info.Phone
        }

        It "Modify user Orgs" {
            $private:info = @{
                PrimaryOrganization = '社長'; Organization = '社長','企画部'
            }
            $private:user = Set-GrnUser -LoginName $UserInfo.LoginName @info -URL $GrnURL -Credential $ValidCred -PassThru -ea Continue            
            $user.LoginName | Should BeExactly $UserInfo.LoginName
            $user.PrimaryOrganization | Should BeExactly $info.PrimaryOrganization
            {$info.Organization -ccontains $user.PrimaryOrganization} | Should Be $true
            $info.Organization.ForEach({
                $user.Organization -ccontains $_
            }) | Should Be $true
        }

        It "Clear user Orgs" {
            $private:info = @{
                Organization = ''
            }
            $private:user = Set-GrnUser -LoginName $UserInfo.LoginName @info -URL $GrnURL -Credential $ValidCred -PassThru -ea Continue -WarningAction SilentlyContinue
            $user.LoginName | Should BeExactly $UserInfo.LoginName
            $info.Organization | Should Be ''
        }

        It "Pipeline from Get-GrnUser" {
            $private:info = @{
                Organization = '総務部','開発部'
                PrimaryOrganization = '情報システム部'
                Position = 30347
            }
            $private:user = Get-GrnUser -LoginName $UserInfo.LoginName -URL $GrnURL -Credential $ValidCred | 
                Set-GrnUser @info -URL $GrnURL -Credential $ValidCred -PassThru -ea Continue            
            $user.LoginName | Should BeExactly $UserInfo.LoginName
            $user.Position | Should BeExactly $info.Position
            $user.PrimaryOrganization | Should BeExactly $info.PrimaryOrganization
            {$info.Organization -ccontains $user.PrimaryOrganization} | Should Be $true
            $info.Organization.ForEach({
                $user.Organization -ccontains $_
            }) | Should Be $true
        }

    }
}
