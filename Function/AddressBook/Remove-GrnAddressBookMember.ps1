
<#
.SYNOPSIS
    ガルーンのアドレス帳からアドレスを削除します
.DESCRIPTION
    ガルーンのアドレス帳からアドレスを削除します
.PARAMETER BookName
    削除対象のアドレス帳の名前を指定します
    BookIdパラメータと同時に使用することはできません
.PARAMETER BookId
    削除対象のアドレス帳のIDを指定します
    BookNameパラメータと同時に使用することはできません
    「個人アドレス帳」に追加する場合は、BookIdに「-1」を指定してください
.PARAMETER CardId
    削除するアドレスのIDを指定します
.PARAMETER URL
    ガルーンのURL
    必ずトップページのURLを指定してください
    例: http://grnserver/cgi-bin/cgi/grn.cgi
.PARAMETER Credential
    ガルーンに接続するための資格情報
.EXAMPLE
    Remove-GrnAddressBookMember -BookName '営業本部' -CardId 2 -URL $URL -Credential $Cred
    Example1: 「営業本部」アドレス帳からID=2のアドレスを削除します
.EXAMPLE
    $Member = Get-GrnAddressBook -BookName '営業本部' -URL $URL -Credential $Cred | Select-Object -ExpandProperty Member
    $Member | ? {$_.DisplayName -eq '鈴木 卓也'} | Remove-GrnAddressBookMember -URL $URL -Credential $Cred
    Example2: 「営業本部」アドレス帳から表示名が「鈴木 卓也」のアドレスを削除します
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

        # アドレスID
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
