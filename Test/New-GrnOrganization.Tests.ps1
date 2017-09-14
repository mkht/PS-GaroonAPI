$moduleRoot = Split-Path $PSScriptRoot -Parent

Get-Module 'PS-GaroonAPI' | Remove-Module -Force
Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force

function Get-RandomName () {
    # なんとなく市区町村ぽい名前をランダム生成する関数
    $MojiList1 = '上中下東西南北京浜丘国寺神和平右左日月火水木金土新古大小海池山河湖黒白一二三'
    $MojiList2 = '市区町村'
    $MojiA = Get-Random -InputObject $MojiList1.ToCharArray() -Count (Get-Random (2,3))
    $MojiB = Get-Random -InputObject $MojiList2.ToCharArray() -Count 1
    $MojiA + $MojiB
}

Describe "Tests of New-GrnOrganization" {
    # 要注意： 本テストを実行するには事前にガルーンデモサイトで「事前設定の反映」を実施しておく必要あり
    $GrnURL = 'https://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
    $ValidCred = New-Object PsCredential "sato", (ConvertTo-SecureString "sato" -asplaintext -force)

    Context '組織の新規作成（組織コード指定あり,親組織なし）' {
        $OrgName = Get-RandomName
        $Code = [System.Web.Security.Membership]::GeneratePassword(9, 0)

        It "実行時にエラーが発生しないか" {
            {$script:Org1 = New-GrnOrganization -OrganizationName $OrgName -OrganizationCode $Code -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop} | Should Not Throw
        }

        It "正しく組織が作成できているか（$OrgName）" {
            $Org1.OrganizationName | Should Be $OrgName
            $Org1.Id | Should Not BeNullOrEmpty
        }

        It "組織コードが指定のものになっているか（$Code)" {
            $Org1.Code | Should Be $Code
        }

        $OrgName = $null
        $Org1 = $null
    }

    Context '組織の新規作成（組織コード指定なし,親組織なし）' {
        $OrgName = Get-RandomName

        It "実行時にエラーが発生しないか" {
            {$script:Org1 = New-GrnOrganization -OrganizationName $OrgName -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop} | Should Not Throw
        }

        It "正しく組織が作成できているか（$OrgName）" {
            $Org1.OrganizationName | Should Be $OrgName
            $Org1.Code | Should Not BeNullOrEmpty
            $Org1.Id | Should Not BeNullOrEmpty
        }

        $OrgName = $null
        $Org1 = $null
    }

    Context '組織の新規作成（パイプライン入力,組織コードあり,親組織あり）' {
        $Parameter = @{
            Name = Get-RandomName
            Code = [System.Web.Security.Membership]::GeneratePassword(8, 0)
            Parent = '経理部'
        }

        It "実行時にエラーが発生しないか" {
            {$script:Org1 = $Parameter | New-GrnOrganization -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop} | Should Not Throw
        }

        It "正しく組織が作成できているか（$Parameter.Name）" {
            $Org1.OrganizationName | Should Be $Parameter.Name
            $Org1.Code | Should Be $Parameter.Code
            $Org1.Id | Should Not BeNullOrEmpty
        }

        $Parameter = $null
        $Org1 = $null
    }

    Context '組織の新規作成（PassThruなし）' {
        $OrgName = Get-RandomName
        $script:Org1 = 'This variable must be null'
        
        It "実行時にエラーが発生しないか" {
            {$script:Org1 = New-GrnOrganization -OrganizationName $OrgName -URL $GrnURL -Credential $ValidCred -ea Stop} | Should Not Throw
        }

        It "なにも出力されないこと確認" {
            $Org1 | Should BeNullOrEmpty
        }

        $OrgName = $null
        $Org1 = $null
    }

    Context '組織の新規作成（PassThruなし）' {
        $OrgName = Get-RandomName
        $script:Org1 = 'This variable must be null'
        
        It "実行時にエラーが発生しないか" {
            {$script:Org1 = New-GrnOrganization -OrganizationName $OrgName -URL $GrnURL -Credential $ValidCred -ea Stop} | Should Not Throw
        }

        It "なにも出力されないこと確認" {
            $Org1 | Should BeNullOrEmpty
        }

        $OrgName = $null
        $Org1 = $null
    }
}
