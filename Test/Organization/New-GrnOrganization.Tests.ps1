#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.3' }

BeforeAll {
    $moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Get-Module 'PS-GaroonAPI' | Remove-Module -Force
    Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force
    Import-Module (Join-Path $moduleRoot './Test/TestUtils/Get-RandomName.psm1') -Force
}

AfterAll {
    Get-Module 'PS-GaroonAPI' | Remove-Module -Force
}

Describe 'Tests of New-GrnOrganization' {
    # 要注意： 本テストを実行するには事前にガルーンデモサイトで「事前設定の反映」を実施しておく必要あり
    BeforeAll {
        $GrnURL = 'https://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
        $ValidCred = New-Object PsCredential 'sato', (ConvertTo-SecureString 'sato' -AsPlainText -Force)

        $BWarn = $global:WarningPreference
        $global:WarningPreference = 'SilentlyContinue'
    }

    Context '組織の新規作成（組織コード指定あり,親組織なし）' {
        BeforeAll {
            $OrgName = Get-RandomName
            $Code = -join ((1..9) | % { Get-Random -input ([char[]]((48..57) + (65..90) + (97..122))) })
        }

        It '実行時にエラーが発生しないか' {
            { $script:Org1 = New-GrnOrganization -OrganizationName $OrgName -OrganizationCode $Code -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop } | Should -Not -Throw
        }

        It "正しく組織が作成できているか（$OrgName）" {
            $Org1.OrganizationName | Should -Be $OrgName
            $Org1.Id | Should -Not -BeNullOrEmpty
        }

        It '親組織は存在しないこと' {
            $Org1.ParentOrganization | Should -BeNullOrEmpty
        }

        It "組織コードが指定のものになっているか（$Code)" {
            $Org1.Code | Should -Be $Code
        }

        AfterAll {
            $script:Org1 = $null
        }
    }

    Context '組織の新規作成（組織コード指定なし,親組織なし）' {
        BeforeAll {
            $OrgName = Get-RandomName
        }

        It '実行時にエラーが発生しないか' {
            { $script:Org1 = New-GrnOrganization -OrganizationName $OrgName -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop } | Should -Not -Throw
        }

        It "正しく組織が作成できているか（$OrgName）" {
            $Org1.OrganizationName | Should -Be $OrgName
            $Org1.Code | Should -Not -BeNullOrEmpty
            $Org1.Id | Should -Not -BeNullOrEmpty
        }

        It '親組織は存在しないこと' {
            $Org1.ParentOrganization | Should -BeNullOrEmpty
        }

        AfterAll {
            $script:Org1 = $null
        }
    }

    Context '組織の新規作成（組織コード指定なし,親組織あり）' {
        BeforeAll {
            $OrgName = Get-RandomName
            $Parent = '企画部'
        }

        It '実行時にエラーが発生しないか' {
            { $script:Org1 = New-GrnOrganization -OrganizationName $OrgName -ParentOrganization $Parent -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop } | Should -Not -Throw
        }

        It "正しく組織が作成できているか（$OrgName）" {
            $Org1.OrganizationName | Should -Be $OrgName
            $Org1.Code | Should -Not -BeNullOrEmpty
            $Org1.Id | Should -Not -BeNullOrEmpty
        }

        It "親組織が正しく設定されているか（$Parent）" {
            $Org1.ParentOrganization | Should -Be $Parent
        }

        AfterAll {
            $script:Org1 = $null
        }
    }

    Context '組織の新規作成（パイプライン入力,組織コードあり,親組織あり）' {
        BeforeAll {
            $Parameter = [PSCustomObject]@{
                Name   = Get-RandomName
                Code   = -join ((1..9) | % { Get-Random -input ([char[]]((48..57) + (65..90) + (97..122))) })
                Parent = '経理部'
            }
        }

        It '実行時にエラーが発生しないか' {
            { $script:Org1 = $Parameter | New-GrnOrganization -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop } | Should -Not -Throw
        }

        It "正しく組織が作成できているか（$($Parameter.Name)）" {
            $Org1.OrganizationName | Should -Be $Parameter.Name
            $Org1.Code | Should -Be $Parameter.Code
            $Org1.Id | Should -Not -BeNullOrEmpty
        }

        AfterAll {
            $Parameter = $null
            $script:Org1 = $null
        }
    }

    Context '組織の新規作成（PassThruなし）' {
        BeforeAll {
            $OrgName = Get-RandomName
            $script:Org1 = 'This variable must be null'
        }

        It '実行時にエラーが発生しないか' {
            { $script:Org1 = New-GrnOrganization -OrganizationName $OrgName -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Not -Throw
        }

        It "正しく組織が作成できているか（$OrgName）" {
            $private:OrgX = Get-GrnOrganization -OrganizationName $OrgName -URL $GrnURL -Credential $ValidCred -NoDetail
            $OrgX.OrganizationName | Should -Be $OrgName
            $OrgX.Code | Should -Not -BeNullOrEmpty
            $OrgX.Id | Should -Not -BeNullOrEmpty
        }

        It 'なにも出力されないこと確認' {
            $Org1 | Should -BeNullOrEmpty
        }

        AfterAll {
            $script:Org1 = $null
        }
    }

    Context '組織コードが重複した場合エラー発生' {
        BeforeAll {
            $OrgName = Get-RandomName
            $OrgCode = 'Executive'  #「役員」と同じ組織コード
            $ExpectErrorMsg = '*[GRN_CMMN_00103]*'
        }

        It '[GRN_CMMN_00103]エラーが発生すること' {
            { New-GrnOrganization -OrganizationName $OrgName -OrganizationCode $OrgCode -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Throw $ExpectErrorMsg
        }

        It "組織が作成されていないこと($OrgName)" {
            $NoOrg = Get-GrnOrganization -OrganizationName $OrgName -NoDetail -URL $GrnURL -Credential $ValidCred -ea Stop
            $NoOrg | Should -BeNullOrEmpty
        }

        AfterAll {
            $script:Org1 = $null
        }
    }

    Context '親組織が見つからない場合エラー発生' {
        BeforeAll {
            $OrgName = Get-RandomName
            $Parent = '存在しない組織'
            $ExpectErrorMsg = '*親組織が見つかりません*'
        }

        It 'エラーが発生すること' {
            { New-GrnOrganization -OrganizationName $OrgName -ParentOrganization $Parent -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Throw $ExpectErrorMsg
        }

        It "組織が作成されていないこと($OrgName)" {
            $NoOrg = Get-GrnOrganization -OrganizationName $OrgName -NoDetail -URL $GrnURL -Credential $ValidCred -ea Stop
            $NoOrg | Should -BeNullOrEmpty
        }

        AfterAll {
            $script:Org1 = $null
        }
    }

    AfterAll {
        $global:WarningPreference = $BWarn
    }
}
