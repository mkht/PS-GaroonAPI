#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.3' }

BeforeAll {
    $moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Get-Module 'PS-GaroonAPI' | Remove-Module -Force
    Import-Module (Join-Path $moduleRoot './PS-GaroonAPI.psd1') -Force
}

AfterAll {
    Get-Module 'PS-GaroonAPI' | Remove-Module -Force
}

Describe 'Get-GrnAddressBook' {
    BeforeAll {
        $GrnURL = 'https://onlinedemo2.cybozu.info/scripts/garoon/grn.exe'
        $ValidCred = New-Object PsCredential 'sato', (ConvertTo-SecureString 'sato' -AsPlainText -Force)
    }

    Context 'Success' {
        It '共有アドレス帳を取得できること（単一ID指定）' {
            $book = Get-GrnAddressBook -BookId 1 -URL $GrnURL -Credential $ValidCred -GetMemberInAddressBook $false
            $book.BookId | Should -Be 1
            $book.Key | Should -Be 'Administration Headquarters'
            $book.Name | Should -Be '経営管理部'
            $book.Type | Should -Be 'cyde'
            $book.Version | Should -Not -BeNullOrEmpty
            $book.CardId | Should -Be (8, 1, 9)
        }

        It '共有アドレス帳を取得できること（複数ID指定）' {
            $book = Get-GrnAddressBook -BookId (1, 3) -URL $GrnURL -Credential $ValidCred -GetMemberInAddressBook $false
            @($book).Count | Should -Be 2
        }

        It '共有アドレス帳を取得できること（単一Name指定）' {
            $book = Get-GrnAddressBook -BookName '経営管理部' -URL $GrnURL -Credential $ValidCred -GetMemberInAddressBook $false
            $book.BookId | Should -Be 1
            $book.Name | Should -Be '経営管理部'
            $book.CardId | Should -Be (8, 1, 9)
        }

        It '共有アドレス帳を取得できること（複数Name指定）' {
            $book = Get-GrnAddressBook -BookName ('経営管理部', '開発本部') -URL $GrnURL -Credential $ValidCred -GetMemberInAddressBook $false
            @($book).Count | Should -Be 2
        }

        It '共有アドレス帳を取得できること（すべて）' {
            $book = Get-GrnAddressBook -URL $GrnURL -Credential $ValidCred -GetMemberInAddressBook $false
            @($book).Count | Should -Be 5
            $book.Name | Should -Contain '個人アドレス帳'
            $book.Name | Should -Contain '経営管理部'
            $book.Name | Should -Contain '情報システム部'
            $book.Name | Should -Contain '営業本部'
            $book.Name | Should -Contain '開発本部'
        }

        It '存在しないアドレス帳を取得できない（ID指定）' {
            $book = Get-GrnAddressBook -BookId 999 -URL $GrnURL -Credential $ValidCred -GetMemberInAddressBook $false
            $book | Should -Be $null
        }

        It '存在しないアドレス帳を取得できない（Name指定）' {
            $book = Get-GrnAddressBook -BookName '存在しない' -URL $GrnURL -Credential $ValidCred -GetMemberInAddressBook $false
            $book | Should -Be $null
        }

        It 'パイプライン入力' {
            $book = '経営管理部' | Get-GrnAddressBook -URL $GrnURL -Credential $ValidCred -GetMemberInAddressBook $false
            $book.Name | Should -Be '経営管理部'
        }

        It '-GetMemberInAddressBook = $true' {
            $book = Get-GrnAddressBook -BookId 1 -URL $GrnURL -Credential $ValidCred -GetMemberInAddressBook $true
            $book.Member | Should -HaveCount 3
        }

        It '-GetMemberInAddressBook = $false' {
            $book = Get-GrnAddressBook -BookId 1 -URL $GrnURL -Credential $ValidCred -GetMemberInAddressBook $false
            $book.Member | Should -HaveCount 0
        }
    }
}
