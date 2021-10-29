<#
.Synopsis
   ガルーンのアドレス帳を取得します
.DESCRIPTION
   ブック名もしくはブックIDをキーにガルーンのアドレス帳を取得します。
   ブック名に何も指定しない場合、閲覧可能なすべてのアドレス帳を取得します。
.EXAMPLE
    Remove-GrnAddressBookMember -BookName '営業本部' -URL 'http://grnserver/grn.cgi' -Credential (Get-Credential)

    BookId  : 3
    Key     : Sales Headquarters
    Name    : 営業本部
    Version : 1191475956
    Type    : cyde
    CardId  : {3, 6, 7, 4}
    Member  : {加藤 美咲, 鈴木 卓也, 音無 結城, 高橋 健太}
#>

function Remove-GrnAddressBookMember {
    [CmdletBinding(DefaultParameterSetName = 'id')]
    Param
    (
        # アドレス帳の名前
        [Parameter(Mandatory = $true, ParameterSetName = 'name')]
        [string]$BookName,

        # アドレス帳ID
        [Parameter(Mandatory = $true, ParameterSetName = 'id', ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [int]$BookId,

        # アドレス帳ID
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [int]$CardId,

        # ガルーンのURL
        [Parameter(Mandatory = $true)]
        [string]$URL,

        # ガルーン管理者の資格情報
        [Parameter(Mandatory = $true)]
        [pscredential]$Credential
    )

    Begin {
        $class = New-Object GaroonAddress @($URL, $Credential) -ErrorAction Stop
        if ($PSCmdlet.ParameterSetName -eq 'name') {
            $targetBook = Get-GrnAddressBook -BookName $BookName -GetMemberInAddressBook $false -URL $URL -Credential $Credential
        }
    }

    Process {
        if ($PSCmdlet.ParameterSetName -eq 'id') {
            $targetBook = Get-GrnAddressBook -BookId $BookId -GetMemberInAddressBook $false -URL $URL -Credential $Credential
        }

        if ($null -eq $targetBook) {
            Write-Error '指定されたアドレス帳が見つかりません'
            return
        }

        try {
            if ($targetBook.CardId -contains $CardId) {
                $class.RemoveSharedCards($CardId, $targetBook.BookId)
            }
            else {
                Write-Warning ('指定されたアドレス帳に CardId={0} のアドレスが存在しません' -f $CardId)
            }
        }
        catch {
            Write-Error -Exception $_.Exception
        }
    }

    End {}
}
