﻿$moduleRoot = Split-Path $PSScriptRoot -Parent

Get-Module 'PS-GaroonAPI' | Remove-Module -Force
Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force

Describe "Tests of Add-GrnOrganizationMember" {
    $GrnURL = 'https://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
    $ValidCred = New-Object PsCredential "sato", (ConvertTo-SecureString "sato" -asplaintext -force)

    Context '組織へユーザ追加(1ユーザ)' {
        $OrgName = '役員'
        $User = 'tanaka'
        $CurrentMembers = ('yamada','sujino','mikami','nomura')

        It "実行時にエラーが発生しないか" {
            {$script:Org1 = Add-GrnOrganizationMember -Name $OrgName -Members $User -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop} | Should Not Throw
        }

        It "正しくユーザが追加されているか" {
            $User -in $Org1.Members | Should Be $true
        }

        It "既存メンバーが維持されているか" {
            @($Org1.Members).Length | Should Be (@($CurrentMembers).Length + 1)
        }

        #繰り返しテスト実行できるようメンバーを元に戻しておく
        Set-GrnOrganization $OrgName -Members $CurrentMembers -URL $GrnURL -Credential $ValidCred -ea SilentlyContinue -wa SilentlyContinue
        $script:Org1 = $null
    }

    Context '組織へユーザ追加(複数ユーザ)' {
        $OrgName = '社長'
        $Users = ('saito', 'miyamoto', 'kato')
        $CurrentMembers = ('hieda', 'takada')

        It "実行時にエラーが発生しないか" {
            {$script:Org1 = Add-GrnOrganizationMember -Name $OrgName -Members $Users -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop} | Should Not Throw
        }

        It "正しくユーザが追加されているか" {
            $Users[0] -in $Org1.Members | Should Be $true
            $Users[1] -in $Org1.Members | Should Be $true
            $Users[2] -in $Org1.Members | Should Be $true
        }

        It "既存メンバーが維持されているか" {
            @($Org1.Members).Length | Should Be (@($CurrentMembers).Length + 3)
        }

        #繰り返しテスト実行できるようメンバーを元に戻しておく
        Set-GrnOrganization $OrgName -Members $CurrentMembers -URL $GrnURL -Credential $ValidCred -ea SilentlyContinue -wa SilentlyContinue
        $script:Org1 = $null
    }

    Context '組織へユーザ追加(パイプライン)' {
        $OrgName = '企画部'
        $Users = ('John', 'brown')
        $CurrentMembers = ('kawakami', 'sakaguchi', 'kamikouya', 'higuma')

        It "実行時にエラーが発生しないか" {
            {$script:Org1 = $Users | Add-GrnOrganizationMember -Name $OrgName -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop} | Should Not Throw
        }

        It "正しくユーザが追加されているか" {
            $Users[0] -in $Org1.Members | Should Be $true
            $Users[1] -in $Org1.Members | Should Be $true
        }

        It "既存メンバーが維持されているか" {
            @($Org1.Members).Length | Should Be (@($CurrentMembers).Length + 2)
        }

        #繰り返しテスト実行できるようメンバーを元に戻しておく
        Set-GrnOrganization $OrgName -Members $CurrentMembers -URL $GrnURL -Credential $ValidCred -ea SilentlyContinue -wa SilentlyContinue
        $script:Org1 = $null
    }

    Context 'Error: 存在しない組織を指定' {
        $OrgName = 'NULL部'
        $User = 'sato'

        It "実行時にエラー発生" {
            {Add-GrnOrganizationMember -Name $OrgName -Members $User -URL $GrnURL -Credential $ValidCred -ea Stop} | Should Throw ('組織 ({0}) が見つかりませんでした' -f $OrgName)
        }
    }
}
