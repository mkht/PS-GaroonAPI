<#
.Synopsis
   ガルーンのアドレス帳を取得します
.DESCRIPTION
   ブック名もしくはブックIDをキーにガルーンのアドレス帳を取得します。
   ブック名に何も指定しない場合、閲覧可能なすべてのアドレス帳を取得します。
.EXAMPLE
    Get-GrnAddressBook -Name '営業本部' -URL 'http://grnserver/grn.cgi' -Credential (Get-Credential)

    BookId  : 3
    Key     : Sales Headquarters
    Name    : 営業本部
    Version : 1191475956
    Type    : cyde
    CardId  : {3, 6, 7, 4}
    Member  : {加藤 美咲, 鈴木 卓也, 音無 結城, 高橋 健太}
#>

function Get-GrnAddressBook {
    [CmdletBinding(DefaultParameterSetName = 'name')]
    Param
    (
        # アドレス帳の名前
        [Parameter(ParameterSetName = 'name', ValueFromPipeline = $true, Position = 0)]
        [Alias('Name')]
        [string[]]$BookName,

        # アドレス帳ID
        [Parameter(ParameterSetName = 'id')]
        [Alias('Id')]
        [ValidateNotNullOrEmpty()]
        [int[]]$BookId,

        # ガルーンのURL
        [Parameter(Mandatory = $true)]
        [string]$URL,

        # ガルーン管理者の資格情報
        [Parameter(Mandatory = $true)]
        [pscredential]$Credential,

        # アドレス帳に登録されているアドレスの詳細情報も取得する
        # Falseの場合アドレスID(CardId)のみ取得し、詳細は取得しない
        [Parameter()]
        [bool]$GetMemberInAddressBook = $true
    )

    Begin {
        $targetBooks = [System.Collections.ArrayList]::new()
        $address = New-Object GaroonAddress @($URL, $Credential) -ErrorAction Stop
        try {
            if ($PSCmdlet.ParameterSetName -eq 'name') {
                [int[]]$BookId = [int[]]::new(0)
                $BookId = $address.GetReadAllowBooks()
                $BookId += -1    #個人アドレス帳のID(-1)を追加
            }

            # $targetBookIds = [System.Collections.Generic.HashSet[int]]::new([int[]]$Id)
            $targetBookIds = $BookId
            foreach ($BookId in $targetBookIds) {
                if ($BookId -le 0) {
                    $targetBooks += $address.GetPersonalBooksById()
                }
                else {
                    $targetBooks += $address.GetSharedBooksById($BookId)
                }
            }
        }
        catch {
            Write-Error -Exception $_.Exception
        }
    }

    Process {
        $returnBooks = [System.Collections.ArrayList]::new()
        if ($PSCmdlet.ParameterSetName -eq 'name') {
            if ($BookName.Count -eq 0) {
                $returnBooks = $targetBooks
            }
            else {
                $returnBooks = ($targetBooks.Where({ $_.Name -in $BookName }))
            }
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'id') {
            $returnBooks = $targetBooks
        }

        foreach ($book in $returnBooks) {
            if ($null -eq $book) {
                continue
            }

            try {
                $result = [ordered]@{
                    BookId  = [int]$book.BookId
                    Key     = [string]$book.key
                    Name    = [string]$book.Name
                    Version = [string]$book.Version
                    Type    = [string]$book.Type
                    CardId  = [int[]]$book.CardId
                    Member  = @()
                }

                if ($GetMemberInAddressBook -and $book.CardId.Count -ge 1) {
                    if ($book.BookId -le 0) {
                        $members = $address.GetPersonalCardsById($book.CardId)
                    }
                    else {
                        $members = $address.GetSharedCardsById($book.CardId, $book.BookId)
                    }

                    if ($members.Count -ge 1) {
                        $result.Member = @($members)
                    }
                }

                [PSCustomObject]$result
            }
            catch {
                Write-Error -Exception $_.Exception
            }
        }
    }

    End {}
}
