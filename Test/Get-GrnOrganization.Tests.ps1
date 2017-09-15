$moduleRoot = Split-Path $PSScriptRoot -Parent

Get-Module 'PS-GaroonAPI' | Remove-Module -Force
Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force

Describe "Tests of Get-GrnOrganization" {
    $script:GrnURL = 'https://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
    $ValidCred = New-Object PsCredential "sato", (ConvertTo-SecureString "sato" -asplaintext -force)

    Context '組織情報を取得する（組織名指定）' {

        It "実行時にエラーが発生しないか" {
            {$script:Org1 = Get-GrnOrganization -OrganizationName '営業部' -URL $GrnURL -Credential $ValidCred -ea Stop} | Should Not Throw
        }

        It "正しい組織名が取得できているか（営業部）" {
            $Org1.OrganizationName | Should Be '営業部'
        }

        It "組織IDが正しく取得できているか（9）" {
            $Org1.Id | Should Be 9
        }

        It "親組織が正しく取得できているか" {
            $Org1.ParentOrganization | Should Be "さいど株式会社"
        }

        It "子組織が正しく取得できているか" {
            $Org1.ChildOrganization.Count | Should Be 2
        }

        It "メンバーが正しく取得できているか" {
            $Org1.Members.Count | Should BeGreaterThan 1
        }
    }

    Context '組織情報を簡易取得する（NoDetail）' {

        It "実行時にエラーが発生しないか" {
            {$script:Org2 = Get-GrnOrganization -OrganizationName '営業部' -NoDetail -URL $GrnURL -Credential $ValidCred -ea Stop} | Should Not Throw
        }

        It "正しい組織名が取得できているか（営業部）" {
            $Org2.OrganizationName | Should Be '営業部'
        }

        It "組織IDが正しく取得できているか（9）" {
            $Org2.Id | Should Be 9
        }

        It "メンバーは取得できていない" {
            $Org2.Members | Should BeNullOrEmpty
        }
    }

    Context '複数の組織情報を取得する' {

        It "実行時にエラーが発生しないか" {
            {$script:Org3 = Get-GrnOrganization -OrganizationName ('営業部', '経理部') -NoDetail -URL $GrnURL -Credential $ValidCred -ea Stop} | Should Not Throw
        }

        It "複数組織が取得できているか" {
            $Org3.Count | Should Be 2
        }

        It "正しい組織名が取得できているか（営業部＆経理部）" {
            $Org3[0].OrganizationName | Should Be '営業部'
            $Org3[1].OrganizationName | Should Be '経理部'
        }
    }

    Context '複数の組織情報を取得する（パイプライン）' {

        It "実行時にエラーが発生しないか" {
            {$script:Org4 = ('営業部', '経理部') | Get-GrnOrganization -URL $GrnURL -Credential $ValidCred -ea Stop} | Should Not Throw
        }

        It '正しい組織名が取得できているか（営業部＆経理部）' {
            $Org4.OrganizationName | Should Be ('営業部', '経理部')
        }

        It '戻り値の型が[PSCustomObject]の配列であること' {
            $Org4 -is [System.object[]] | Should Be $true
            $Org4[0] -is [PSCustomObject] | Should Be $true
            $Org4[1] -is [PSCustomObject] | Should Be $true
        }
    }

    Context '組織情報を取得する(ワイルドカード検索)' {

        It "実行時にエラーが発生しないか" {
            {$script:Org5 = Get-GrnOrganization -OrganizationName '*部' -SearchMode Like -NoDetail -URL $GrnURL -Credential $ValidCred -ea Stop} | Should Not Throw
        }

        It "複数組織が取得できているか" {
            $Org5.Count | Should BeGreaterThan 1
        }

        It "マッチする組織が取得できているか" {
            ($Org5.OrganizationName | ? {$_ -like '*部'}).Count | Should Be $Org5.OrganizationName.Count
        }

        It "マッチしない組織が含まれていないか" {
            $Org5.OrganizationName | ? {$_ -notlike '*部'} | Should BeNullOrEmpty
        }
    }

    Context '組織情報を取得する(正規表現検索)' {

        It "実行時にエラーが発生しないか" {
            {$script:Org6 = Get-GrnOrganization -OrganizationName '第\d営業' -SearchMode RegExp -NoDetail -URL $GrnURL -Credential $ValidCred -ea Stop} | Should Not Throw
        }

        It "複数組織が取得できているか" {
            $Org6.Count | Should Be 2
        }

        It "マッチする組織が取得できているか" {
            ($Org6.OrganizationName | ? {$_ -match '第\d営業'}).Count | Should Be $Org6.OrganizationName.Count
        }

        It "マッチしない組織が含まれていないか" {
            $Org6.OrganizationName | ? {$_ -notmatch '第\d営業'} | Should BeNullOrEmpty
        }
    }

    Context '存在しない組織情報は取得できない' {

        It "実行時にエラーが発生しないか" {
            {$script:Org7 = Get-GrnOrganization -OrganizationName 'ほげほげ部' -URL $GrnURL -Credential $ValidCred -ea Stop} | Should Not Throw
        }

        It '$nullを返しているか' {
            $Org7 | Should BeNullOrEmpty
        }
    }

}
