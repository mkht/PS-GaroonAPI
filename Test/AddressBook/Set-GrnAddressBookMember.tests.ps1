#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.3' }

BeforeAll {
    $moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Get-Module 'PS-GaroonAPI' | Remove-Module -Force
    Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force
}

AfterAll {
    Get-Module 'PS-GaroonAPI' | Remove-Module -Force
}

Describe 'Set-GrnAddressBookMember' {
    BeforeAll {
        $GrnURL = 'https://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
        $ValidCred = New-Object PsCredential 'sato', (ConvertTo-SecureString 'sato' -AsPlainText -Force)
    }

    Context 'Success' {
        It '共有アドレス帳のアドレス編集（ID指定）' {
            $Card = @{
                CardId         = 7
                DisplayName    = '音有 結城'
                FamilyNameKana = 'おとあり'
                GivenNameKana  = 'ゆうき'
                CompanyName    = 'サイボウズ（株）'
            }
            $ret = Set-GrnAddressBookMember @Card -BookId 3 -URL $GrnURL -Credential $ValidCred -PassThru
            $ret.BookId | Should -Be 3
            $ret.CardId | Should -Be 7
            $ret.DisplayName | Should -Be '音有 結城'
            $ret.FamilyName | Should -BeNullOrEmpty
            $ret.GivenName | Should -BeNullOrEmpty
            $ret.FamilyNameKana | Should -Be 'おとあり'
            $ret.GivenNameKana | Should -Be 'ゆうき'
            $ret.CompanyName | Should -Be 'サイボウズ（株）'
        }

        It '共有アドレス帳のアドレス編集（Name指定）' {
            $Card = @{
                CardId      = 3
                DisplayName = 'ポール美咲'
            }
            $ret = Set-GrnAddressBookMember @Card -BookName '営業本部' -URL $GrnURL -Credential $ValidCred -PassThru
            $ret.BookId | Should -Be 3
            $ret.CardId | Should -Be 3
            $ret.DisplayName | Should -Be 'ポール美咲'
        }

        It '個人アドレス帳のアドレス編集' {
            $Card = @{
                CardId      = 8
                DisplayName = '吉田 茂'
            }
            $ret = Set-GrnAddressBookMember @Card -BookId -1 -URL $GrnURL -Credential $ValidCred -PassThru
            $ret.CardId | Should -Be 8
            $ret.DisplayName | Should -Be '吉田 茂'
        }

        It 'PassThruなしの場合は戻り値なし' {
            $Card = @{
                CardId      = 6
                DisplayName = 'Takuya Suzuki'
            }
            $ret = Set-GrnAddressBookMember @Card -BookId 3 -URL $GrnURL -Credential $ValidCred
            $ret | Should -Be $null
        }

        It 'パイプライン入力' {
            $Card = [pscustomobject]@{
                CardId      = 3
                DisplayName = 'ヤマダ工務店'
                ZipCode     = '123-4567'
                Address     = '埼玉県さいたま市見沼2-3'
            }
            $ret = $Card | Set-GrnAddressBookMember -BookId 3 -URL $GrnURL -Credential $ValidCred -PassThru
            $ret.BookId | Should -Be 3
            $ret.CardId | Should -Be 3
            $ret.DisplayName | Should -Be 'ヤマダ工務店'
        }

        It '存在しないCardIdを指定(-Forceあり)' {
            $Card = @{
                CardId      = 99
                DisplayName = '麺屋三郎'
            }
            Mock Add-GrnAddressBookMember { return $true } -ModuleName PS-GaroonAPI
            { Set-GrnAddressBookMember @Card -BookId 3 -URL $GrnURL -Credential $ValidCred -Force -ea Stop } | Should -Not -Throw
            Should -Invoke Add-GrnAddressBookMember -Times 1 -Exactly -ModuleName PS-GaroonAPI
        }
    }

    Context 'Error' {
        It '存在しないアドレス帳を指定' {
            $Card = @{
                CardId      = 3
                DisplayName = 'Thunder Pacha'
            }
            { Set-GrnAddressBookMember @Card -BookId 999 -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Throw '指定されたアドレス帳が見つかりません'
        }

        It '存在しないCardIdを指定(-Forceなし)' {
            $Card = @{
                CardId      = 99
                DisplayName = 'ホームセンター'
            }
            { Set-GrnAddressBookMember @Card -BookId 3 -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Throw '*アドレスが存在しません*'
        }
    }
}
