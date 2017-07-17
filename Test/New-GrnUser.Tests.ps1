$moduleRoot = Split-Path $PSScriptRoot -Parent

Get-Module 'PS-GaroonAPI' | Remove-Module -Force
Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force

Describe "Tests of New-GrnUser" {
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

            {New-GrnUser @DummyInfo -URL $WrongURL -Credential $ValidCred -ea Stop} | Should throw '404'
        }

        It "If you use invalid credential, shoud throw [FW00007] error" {
            $local:InvalidCred = New-Object PsCredential "hoge",(ConvertTo-SecureString "hoge" -asplaintext -force)
            
            {New-GrnUser @DummyInfo -URL $GrnURL -Credential $InvalidCred -ea Stop} | Should throw '[FW00007]'
        }

        It "If try to create duplicate user, shoud throw" {
            
            $local:DuplicateUser = @{
                LoginName = 'sato'
                Password = 'sato'
                DisplayName = '佐藤 昇'
            }
            
            {New-GrnUser @DuplicateUser -URL $GrnURL -Credential $ValidCred -ea Stop} | Should throw '既に存在します'
        }

        It "If specified NON existing Organization name, shoud throw" {
            
            $local:DuplicateUser = $DummyInfo
            $DuplicateUser.PrimaryOrganization = 'ほにゃららら'
            $DuplicateUser.Organization = '役員'
            
            {New-GrnUser @DuplicateUser -URL $GrnURL -Credential $ValidCred -ea Stop} | Should throw '存在しない組織名が含まれています'
        }
    }

    Context 'Success' {
        $local:Gen_Id = {
            $chars = 'abcdefghkmnprstuvwxyzABCDEFGHKLMNPRSTUVWXYZ123456789'
            $random = 1..8 | ForEach-Object { Get-Random -Maximum $chars.length }
            [string]($chars[$random] -join '')
        }

        It "Create new user with orgs" {
            $script:UserInfo = @{
                LoginName = & $Gen_Id
                DisplayName = '手巣戸 太郎'
                Password = 'P@ssw0rd'
                Kana = 'てすと たろう'
                PrimaryOrganization = '役員'
                Organization = "総務部","経理部"
            }

            $private:user = New-GrnUser @UserInfo -URL $GrnURL -Credential $ValidCred -PassThru -ea Continue            
            $user.LoginName | Should BeExactly $UserInfo.LoginName
            $user.DisplayName | Should BeExactly $UserInfo.DisplayName
            $user.PrimaryOrganization | Should BeExactly $UserInfo.PrimaryOrganization
            {$UserInfo.Organization -ccontains $user.PrimaryOrganization} | Should Be $true
            $UserInfo.Organization.ForEach({
                $user.Organization -ccontains $_
            }) | Should Be $true
        }

        It "Create new user without orgs" {
            $script:UserInfo = @{
                LoginName = & $Gen_Id
                DisplayName = '手巣戸 花子'
                Password = 'P@ssw0rd'
                PrimaryOrganization = ''
                Organization = ''
            }

            $private:user = New-GrnUser @UserInfo -URL $GrnURL -Credential $ValidCred -PassThru -ea Continue            
            $user.LoginName | Should BeExactly $UserInfo.LoginName
            $user.DisplayName | Should BeExactly $UserInfo.DisplayName
            $UserInfo.Organization | Should Be ''
        }

        AfterEach{
            if($UserInfo){
                Set-GrnUser -LoginName $UserInfo.LoginName -Invalid $true -URL $GrnURL -Credential $ValidCred -ea Continue
            }
        }
    }
}
