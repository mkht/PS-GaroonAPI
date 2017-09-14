$moduleRoot = Split-Path $PSScriptRoot -Parent

Get-Module 'PS-GaroonAPI' | Remove-Module -Force
Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force

function Get-RandomName () {
    # なんとなく市区町村ぽい名前をランダム生成する関数
    $MojiList1 = '上中下東西南北京浜丘国寺神和平右左日月火水木金土新古大小海池山河湖黒白一二三'
    $MojiList2 = '市区町村'
    $MojiA = Get-Random -InputObject $MojiList1.ToCharArray() -Count (Get-Random (2, 3))
    $MojiB = Get-Random -InputObject $MojiList2.ToCharArray() -Count 1
    -join ($MojiA + $MojiB)
}

Describe "Tests of New-GrnOrganization" {
    # 要注意： 本テストを実行するには事前にガルーンデモサイトで「事前設定の反映」を実施しておく必要あり
    $GrnURL = 'https://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
    $ValidCred = New-Object PsCredential "sato", (ConvertTo-SecureString "sato" -asplaintext -force)

    $BWarn = $global:WarningPreference
    $global:WarningPreference = 'SilentlyContinue'

    Context '組織の新規作成（組織コード指定あり,親組織なし）' {
        $OrgName = Get-RandomName
        $Code = -join ((1..9) | % {Get-Random -input ([char[]]((48..57) + (65..90) + (97..122)))})

        It "実行時にエラーが発生しないか" {
            {$script:Org1 = New-GrnOrganization -OrganizationName $OrgName -OrganizationCode $Code -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop} | Should Not Throw
        }

        It "正しく組織が作成できているか（$OrgName）" {
            $Org1.OrganizationName | Should Be $OrgName
            $Org1.Id | Should Not BeNullOrEmpty
        }

        It "親組織は存在しないこと" {
            $Org1.ParentOrganization | Should BeNullOrEmpty
        }

        It "組織コードが指定のものになっているか（$Code)" {
            $Org1.Code | Should Be $Code
        }

        $script:Org1 = $null
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

        It "親組織は存在しないこと" {
            $Org1.ParentOrganization | Should BeNullOrEmpty
        }

        $script:Org1 = $null
    }

    Context '組織の新規作成（組織コード指定なし,親組織あり）' {
        $OrgName = Get-RandomName
        $Parent = '企画部'

        It "実行時にエラーが発生しないか" {
            {$script:Org1 = New-GrnOrganization -OrganizationName $OrgName -ParentOrganization $Parent -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop} | Should Not Throw
        }

        It "正しく組織が作成できているか（$OrgName）" {
            $Org1.OrganizationName | Should Be $OrgName
            $Org1.Code | Should Not BeNullOrEmpty
            $Org1.Id | Should Not BeNullOrEmpty
        }

        It "親組織が正しく設定されているか（$Parent）" {
            $Org1.ParentOrganization | Should Be $Parent
        }

        $script:Org1 = $null
    }

    Context '組織の新規作成（パイプライン入力,組織コードあり,親組織あり）' {
        $Parameter = [PSCustomObject]@{
            Name   = Get-RandomName
            Code   = -join ((1..9) | % {Get-Random -input ([char[]]((48..57) + (65..90) + (97..122)))})
            Parent = '経理部'
        }

        It "実行時にエラーが発生しないか" {
            {$script:Org1 = $Parameter | New-GrnOrganization -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop} | Should Not Throw
        }

        It "正しく組織が作成できているか（$($Parameter.Name)）" {
            $Org1.OrganizationName | Should Be $Parameter.Name
            $Org1.Code | Should Be $Parameter.Code
            $Org1.Id | Should Not BeNullOrEmpty
        }

        $Parameter = $null
        $script:Org1 = $null
    }

    Context '組織の新規作成（PassThruなし）' {
        $OrgName = Get-RandomName
        $script:Org1 = 'This variable must be null'

        It "実行時にエラーが発生しないか" {
            {$script:Org1 = New-GrnOrganization -OrganizationName $OrgName -URL $GrnURL -Credential $ValidCred -ea Stop} | Should Not Throw
        }

        It "正しく組織が作成できているか（$OrgName）" {
            $private:OrgX = Get-GrnOrganization -OrganizationName $OrgName -URL $GrnURL -Credential $ValidCred -NoDetail
            $OrgX.OrganizationName | Should Be $OrgName
            $OrgX.Code | Should Not BeNullOrEmpty
            $OrgX.Id | Should Not BeNullOrEmpty
        }

        It "なにも出力されないこと確認" {
            $Org1 | Should BeNullOrEmpty
        }

        $script:Org1 = $null
    }

    Context '組織コードが重複した場合エラー発生' {
        $OrgName = Get-RandomName
        $OrgCode = 'Executive'  #「役員」と同じ組織コード
        $ExpectErrorMsg = '[ERROR][GRN_CMMN_00103] すでに存在する組織コードの組織を指定しています。'

        It "[GRN_CMMN_00103]エラーが発生すること" {
            {New-GrnOrganization -OrganizationName $OrgName -OrganizationCode $OrgCode -URL $GrnURL -Credential $ValidCred -ea Stop} | Should Throw $ExpectErrorMsg
        }

        It "組織が作成されていないこと($OrgName)" {
            $NoOrg = Get-GrnOrganization -OrganizationName $OrgName -NoDetail -URL $GrnURL -Credential $ValidCred -ea Stop
            $NoOrg | Should BeNullOrEmpty
        }

        $script:Org1 = $null
    }

    Context '親組織が見つからない場合エラー発生' {
        $OrgName = Get-RandomName
        $Parent = '存在しない組織'
        $ExpectErrorMsg = '親組織が見つかりません'

        It "エラーが発生すること" {
            {New-GrnOrganization -OrganizationName $OrgName -ParentOrganization $Parent -URL $GrnURL -Credential $ValidCred -ea Stop} | Should Throw $ExpectErrorMsg
        }

        It "組織が作成されていないこと($OrgName)" {
            $NoOrg = Get-GrnOrganization -OrganizationName $OrgName -NoDetail -URL $GrnURL -Credential $ValidCred -ea Stop
            $NoOrg | Should BeNullOrEmpty
        }

        $script:Org1 = $null
    }

    $global:WarningPreference = $BWarn
}
