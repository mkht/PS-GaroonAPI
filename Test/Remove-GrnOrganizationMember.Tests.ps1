$moduleRoot = Split-Path $PSScriptRoot -Parent

Get-Module 'PS-GaroonAPI' | Remove-Module -Force
Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force

Describe 'Tests of Remove-GrnOrganizationMember' {

    BeforeAll {
        $script:GrnURL = 'https://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
        $script:ValidCred = New-Object PsCredential 'sato', (ConvertTo-SecureString 'sato' -AsPlainText -Force)
        $script:OrgName = '役員'
        $script:CurrentMembers = ('yamada', 'sujino', 'mikami', 'nomura')

        Set-GrnOrganization $OrgName -Members $CurrentMembers -URL $GrnURL -Credential $ValidCred -ea Stop -wa SilentlyContinue
        $script:Org1 = $null
    }

    AfterEach {
        #繰り返しテスト実行できるようメンバーを元に戻す
        Set-GrnOrganization $OrgName -Members $CurrentMembers -URL $GrnURL -Credential $ValidCred -ea Stop -wa SilentlyContinue
        $script:Org1 = $null
    }

    It 'PassThruスイッチ無しの場合何も出力しない' {
        $OneUser = 'yamada'

        # 何も出力しない
        Remove-GrnOrganizationMember -Name $OrgName -Members $OneUser -URL $GrnURL -Credential $ValidCred -ea Stop | Should BeNullOrEmpty
    }

    It 'PassThruスイッチ有りの場合はオブジェクト出力' {
        $OneUser = 'sujino'

        # 出力あり
        { Remove-GrnOrganizationMember -Name $OrgName -Members $OneUser -URL $GrnURL -Credential $ValidCred -ea Stop -PassThru } | Should Not BeNullOrEmpty
    }

    It '組織からユーザ削除(1ユーザ)' {
        $OneUser = 'mikami'

        # ユーザ削除実行
        $Org1 = Remove-GrnOrganizationMember -Name $OrgName -Members $OneUser -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop

        # 対象ユーザが削除されているか
        $User -in $Org1.Members | Should Be $false

        # メンバー数が減っているか
        @($Org1.Members).Length | Should Be (@($CurrentMembers).Length - 1)
    }

    It '組織からユーザ削除(複数ユーザ)' {
        $TwoUsers = ('sujino', 'nomura')

        # ユーザ削除実行
        $Org1 = Remove-GrnOrganizationMember -Name $OrgName -Members $TwoUsers -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop

        # 対象ユーザが削除されているか
        $TwoUsers[0] -in $Org1.Members | Should Be $false
        $TwoUsers[1] -in $Org1.Members | Should Be $false

        # メンバー数が減っているか
        @($Org1.Members).Length | Should Be (@($CurrentMembers).Length - 2)
    }

    It '組織からユーザ削除(パイプライン)' {
        $TwoUsers = ('mikami', 'nomura')

        # ユーザ削除実行
        $Org1 = $TwoUsers | Remove-GrnOrganizationMember -Name $OrgName -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop

        # 対象ユーザが削除されているか
        $TwoUsers[0] -in $Org1.Members | Should Be $false
        $TwoUsers[1] -in $Org1.Members | Should Be $false

        # メンバー数が減っているか
        @($Org1.Members).Length | Should Be (@($CurrentMembers).Length - 2)

    }

    It '元々組織のメンバーでないユーザを消そうとしてもエラー・警告なし' {
        $NoMember = 'sato'

        # エラー・警告とも発生しないこと
        { $script:Org1 = Remove-GrnOrganizationMember -Name $OrgName -Members $NoMember -URL $GrnURL -Credential $ValidCred -PassThru -ea Stop -wa Stop } | Should Not Throw

        # 組織内のメンバーに変化がないこと
        Compare-Object $CurrentMembers $Org1.Members -PassThru | Should BeNullOrEmpty
        @($Org1.Members).Length -eq @($CurrentMembers).Length | Should Be $true
    }

    It 'ガルーンに存在しないユーザを消そうとした場合は警告メッセージ' {
        $NullMember = 'NULLNULL'
        $ExpectMsg = ('指定されたログイン名のユーザ({0})が見つかりません' -f $NullMember)

        # 警告が発生すること
        $WarnMsg = Remove-GrnOrganizationMember -Name $OrgName -Members $NullMember -URL $GrnURL -Credential $ValidCred -ea Stop 3>&1
        $WarnMsg.Message | Should Be $ExpectMsg
    }

    It 'Error: 存在しない組織を指定した場合は例外発生' {
        $OrgName = 'NULL部'
        $OneUser = 'sato'
        $ExpectMsg = ('組織 ({0}) が見つかりませんでした' -f $OrgName)

        # 実行時にエラー発生
        { Remove-GrnOrganizationMember -Name $OrgName -Members $OneUser -URL $GrnURL -Credential $ValidCred -ea Stop } | Should Throw $ExpectMsg
    }
}
