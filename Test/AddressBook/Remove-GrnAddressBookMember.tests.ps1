#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.3' }

BeforeAll {
    $moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Get-Module 'PS-GaroonAPI' | Remove-Module -Force
    Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force
}

AfterAll {
    Get-Module 'PS-GaroonAPI' | Remove-Module -Force
}

Describe 'Remove-GrnAddressBookMember' {
    BeforeAll {
        $GrnURL = 'https://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
        $ValidCred = New-Object PsCredential 'sato', (ConvertTo-SecureString 'sato' -AsPlainText -Force)
        $script:TestCardId = 4
        $script:TestBookId = 3
    }

    BeforeEach {
        $warn = $null
        $script:TestCardId = (Set-GrnAddressBookMember -BookId $TestBookId -CardId $TestCardId -DisplayName 'test' -PassThru -Force -URL $GrnURL -Credential $ValidCred).CardId
    }

    Context 'Success' {
        It '共有アドレス帳のアドレス削除（ID指定）' {
            $ret = Remove-GrnAddressBookMember -BookId $TestBookId -CardId $TestCardId -URL $GrnURL -Credential $ValidCred -WarningVariable warn
            $warn | Should -Be $null
            $ret | Should -Be $null
            $member = Get-GrnAddressBook -BookId $TestBookId -URL $GrnURL -Credential $ValidCred -GetMemberInAddressBook $false
            $member.CardId | Should -Not -Contain $TestCardId
        }

        It '共有アドレス帳のアドレス削除（Name指定）' {
            $ret = Remove-GrnAddressBookMember -BookName '営業本部' -CardId $TestCardId -URL $GrnURL -Credential $ValidCred -WarningVariable warn
            $warn | Should -Be $null
            $ret | Should -Be $null
            $member = Get-GrnAddressBook -BookId $TestBookId -URL $GrnURL -Credential $ValidCred -GetMemberInAddressBook $false
            $member.CardId | Should -Not -Contain $TestCardId
        }

        It 'パイプライン入力' {
            $Card = [pscustomobject]@{
                BookId = $TestBookId
                CardId = $TestCardId
            }
            $ret = $Card | Remove-GrnAddressBookMember -URL $GrnURL -Credential $ValidCred -WarningVariable warn
            $warn | Should -Be $null
            $ret | Should -Be $null
            $member = Get-GrnAddressBook -BookId $TestBookId -URL $GrnURL -Credential $ValidCred -GetMemberInAddressBook $false
            $member.CardId | Should -Not -Contain $TestCardId
        }
    }

    Context 'Error' {
        It '存在しないアドレス帳を指定' {
            { Remove-GrnAddressBookMember -CardId $TestCardId -BookId 999 -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Throw '指定されたアドレス帳が見つかりません'
        }

        It '存在しないCardIdを指定' {
            Remove-GrnAddressBookMember -CardId 999 -BookId $TestBookId -URL $GrnURL -Credential $ValidCred -ea Stop -WarningVariable warn
            $warn | Should -BeLike '*アドレスが存在しません*'
        }
    }
}
