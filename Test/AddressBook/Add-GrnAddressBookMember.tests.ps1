$moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent

Get-Module 'PS-GaroonAPI' | Remove-Module -Force
Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force

InModuleScope PS-GaroonAPI {
    Describe 'Add-GrnAddressBookMember' {
        BeforeAll {
            $GrnURL = 'https://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
            $ValidCred = New-Object PsCredential 'sato', (ConvertTo-SecureString 'sato' -AsPlainText -Force)
        }

        Context 'Success' {
            It '共有アドレス帳にアドレス追加（ID指定）' {
                $Card = @{
                    DisplayName    = '佐藤　健二'
                    FamilyName     = '佐藤'
                    GivenName      = '健二'
                    FamilyNameKana = 'さとう'
                    GivenNameKana  = 'けんじ'
                    CompanyName    = '株式会社テスト'
                    CompanyReading = 'かぶしきがいしゃてすと'
                    Section        = '広報部'
                    ZipCode        = '123-4567'
                    Address        = '東京都港区芝公園'
                    CompanyPhone   = '03-1234-5678'
                    CompanyFax     = '03-1234-5679'
                    Link           = 'http://www.example.com/'
                    Post           = '部長'
                    Phone          = '03-1234-5678'
                    Email          = 'foo-bar@example.com'
                    Description    = 'テスト'
                }
                $ret = Add-GrnAddressBookMember @Card -BookId 2 -URL $GrnURL -Credential $ValidCred -PassThru
                $ret.BookId | Should -Be 2
                $ret.CardId | Should -BeGreaterThan 0
                $ret.DisplayName | Should -Be '佐藤　健二'
                $ret.FamilyName | Should -Be '佐藤'
                $ret.GivenName | Should -Be '健二'
                $ret.FamilyNameKana | Should -Be 'さとう'
                $ret.GivenNameKana | Should -Be 'けんじ'
                $ret.CompanyName | Should -Be '株式会社テスト'
                $ret.CompanyReading | Should -Be 'かぶしきがいしゃてすと'
                $ret.Section | Should -Be '広報部'
                $ret.ZipCode | Should -Be '123-4567'
                $ret.Address | Should -Be '東京都港区芝公園'
                $ret.CompanyPhone | Should -Be '03-1234-5678'
                $ret.CompanyFax | Should -Be '03-1234-5679'
                $ret.URL | Should -Be 'http://www.example.com/'
                $ret.Post | Should -Be '部長'
                $ret.Phone | Should -Be '03-1234-5678'
                $ret.Email | Should -Be 'foo-bar@example.com'
                $ret.Description | Should -Be 'テスト'
            }

            It '共有アドレス帳にアドレス追加（Name指定）' {
                $Card = @{
                    DisplayName = '田中千代子'
                }
                $ret = Add-GrnAddressBookMember @Card -BookName '情報システム部' -URL $GrnURL -Credential $ValidCred -PassThru
                $ret.BookId | Should -Be 2
                $ret.CardId | Should -BeGreaterThan 0
                $ret.DisplayName | Should -Be '田中千代子'
            }

            It '個人アドレス帳にアドレス追加' {
                $Card = @{
                    DisplayName = 'عَبْدُ الْقُدُّوسِ'
                }
                $ret = Add-GrnAddressBookMember @Card -BookId -1 -URL $GrnURL -Credential $ValidCred -PassThru
                $ret.CardId | Should -BeGreaterThan 0
                $ret.DisplayName | Should -Be 'عَبْدُ الْقُدُّوسِ'
            }

            It 'PassThruなしの場合は戻り値なし' {
                $Card = @{
                    DisplayName = 'Alex Thompson'
                }
                $ret = Add-GrnAddressBookMember @Card -BookId 2 -URL $GrnURL -Credential $ValidCred
                $ret | Should -Be $null
            }

            It 'パイプライン入力' {
                $Card = [pscustomobject]@{
                    DisplayName = 'ヤマダ工務店'
                    ZipCode     = '123-4567'
                    Address     = '埼玉県さいたま市見沼2-3'
                }
                $ret = $Card | Add-GrnAddressBookMember -BookId 2 -URL $GrnURL -Credential $ValidCred -PassThru
                $ret.BookId | Should -Be 2
                $ret.CardId | Should -BeGreaterThan 0
                $ret.DisplayName | Should -Be 'ヤマダ工務店'
            }
        }

        Context 'Error' {
            It '存在しないアドレス帳を指定' {
                $Card = @{
                    DisplayName = 'Thunder Pacha'
                }
                { Add-GrnAddressBookMember @Card -BookId 999 -URL $GrnURL -Credential $ValidCred -ea Stop } | Should -Throw '指定されたアドレス帳が見つかりません'
            }
        }
    }
}
