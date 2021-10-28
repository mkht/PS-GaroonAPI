$moduleRoot = Split-Path $PSScriptRoot -Parent

Get-Module 'PS-GaroonAPI' | Remove-Module -Force
Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force

Describe 'Tests of Set-GrnOrganization' {
    # 要注意： 本テストを実行するには事前にガルーンデモサイトで「事前設定の反映」を実施しておく必要あり
    $GrnURL = 'https://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
    $ValidCred = New-Object PsCredential 'sato', (ConvertTo-SecureString 'sato' -AsPlainText -Force)

    Context '組織名の変更' {
        $OrgName = '第1営業グループ'
        $New = 'XXX_第1営業グループ'

        It '実行時にエラーが発生しないか' {
            { $script:Org1 = Set-GrnOrganization -Name $OrgName -NewName $New -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop } | Should Not Throw
        }

        It '正しく組織名が変更されているか' {
            $Org1.OrganizationName | Should Be $New
        }

        #繰り返しテスト実行できるよう組織名を元に戻しておく
        Set-GrnOrganization $New -NewName $OrgName -URL $GrnURL -Credential $ValidCred -ea SilentlyContinue -wa SilentlyContinue
        $script:Org1 = $null
    }

    Context '親組織の変更' {
        $OrgName = '営業部'
        $CurrentParent = 'さいど株式会社'
        $NewParent = '企画部'

        It '実行時にエラーが発生しないか' {
            { $script:Org1 = Set-GrnOrganization -Name $OrgName -Parent $NewParent -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop } | Should Not Throw
        }

        It '正しく親組織が変わっているか' {
            $Org1.ParentOrganization | Should Be $NewParent
        }

        #繰り返しテスト実行できるよう親組織を元に戻しておく
        Set-GrnOrganization $OrgName -Parent $CurrentParent -URL $GrnURL -Credential $ValidCred -ea SilentlyContinue -wa SilentlyContinue
        $script:Org1 = $null
    }

    Context '新しい親組織が現在の親組織と同じ' {
        $OrgName = '役員'
        $Parent = 'さいど株式会社'

        It '実行時にエラー発生しない' {
            { Set-GrnOrganization -Name $OrgName -Parent $Parent -URL $GrnURL -Credential $ValidCred -ea Stop 6>$null } | Should Not Throw
        }
    }

    Context 'メンバーの変更' {
        $OrgName = '情報システム部'
        $NewMember = ('sakaguchi', 'higuma', 'brown', 'yamada')
        $CurrentMember = ('yamada', 'sato', 'John', 'ito', 'matsuda')

        It '実行時にエラーが発生しないか' {
            { $script:Org1 = Set-GrnOrganization -Name $OrgName -Members $NewMember -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop } | Should Not Throw
        }

        It '正しくメンバーが変わっているか' {
            @($Org1.Members.Count) | Should Be 4
            $Org1.Members -ccontains $NewMember[0] | Should Be $true
            $Org1.Members -ccontains $NewMember[1] | Should Be $true
            $Org1.Members -ccontains $NewMember[2] | Should Be $true
            $Org1.Members -ccontains $NewMember[3] | Should Be $true
        }

        #繰り返しテスト実行できるようメンバーを元に戻しておく
        Set-GrnOrganization $OrgName -Members $CurrentMember -URL $GrnURL -Credential $ValidCred -ea SilentlyContinue -wa SilentlyContinue
        $script:Org1 = $null
    }

    Context 'メンバーの変更(メンバー全削除)' {
        $OrgName = '開発部'
        $NewMember = @()
        $CurrentMember = ('John', 'kobayashi', 'furukawa', 'davis', 'li')

        It '実行時にエラーが発生しないか' {
            { $script:Org1 = Set-GrnOrganization -Name $OrgName -Members $NewMember -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop } | Should Not Throw
        }

        It '正しくメンバーが消えているか' {
            $Org1.Members | Should BeNullOrEmpty
        }

        #繰り返しテスト実行できるようメンバーを元に戻しておく
        Set-GrnOrganization $OrgName -Members $CurrentMember -URL $GrnURL -Credential $ValidCred -ea SilentlyContinue -wa SilentlyContinue
        $script:Org1 = $null
    }

    Context 'Error: 存在しない組織は変更できない' {
        $OrgName = 'NULL部'

        It '実行時にエラー発生' {
            { Set-GrnOrganization -Name $OrgName -Members @() -URL $GrnURL -Credential $ValidCred -ea Stop } | Should Throw '見つかりませんでした'
        }
    }

    Context 'Error: 新しい組織名に既に存在する組織名を指定' {
        $OrgName = '役員'
        $NewName = '社長'

        It '実行時にエラー発生' {
            { Set-GrnOrganization -Name $OrgName -NewName $NewName -URL $GrnURL -Credential $ValidCred -ea Stop } | Should Throw '組織が既に存在する'
        }
    }

    Context 'Error: 親組織名に存在しない組織を指定' {
        $OrgName = '役員'
        $Parent = 'NULL部'

        It '実行時にエラー発生' {
            { Set-GrnOrganization -Name $OrgName -Parent $Parent -URL $GrnURL -Credential $ValidCred -ea Stop } | Should Throw '新しい親組織'
        }
    }
}
