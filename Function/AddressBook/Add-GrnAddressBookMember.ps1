<#
.SYNOPSIS
    ガルーンのアドレス帳にアドレスを追加します
.DESCRIPTION
    ガルーンのアドレス帳にアドレスを追加します
.PARAMETER BookName
    追加対象のアドレス帳の名前を指定します
    BookIdパラメータと同時に使用することはできません
.PARAMETER BookId
    追加対象のアドレス帳のIDを指定します
    BookNameパラメータと同時に使用することはできません
    「個人アドレス帳」に追加する場合は、BookIdに「-1」を指定してください
.PARAMETER DisplayName
    追加するアドレスの表示名
    必須パラメータです
.PARAMETER FamilyName
    追加するアドレスの姓
.PARAMETER GivenName
    追加するアドレスの名
.PARAMETER FamilyNameKana
    追加するアドレスの姓（よみ）
.PARAMETER GivenNameKana
    追加するアドレスの名（よみ）
.PARAMETER CompanyName
    追加するアドレスの会社名
.PARAMETER CompanyReading
    追加するアドレスの会社名（よみ）
.PARAMETER Section
    追加するアドレスの部署名
.PARAMETER ZipCode
    追加するアドレスの郵便番号
.PARAMETER Address
    追加するアドレスの住所
.PARAMETER Map
    追加するアドレスの地図URL
.PARAMETER CompanyPhone
    追加するアドレスの会社電話番号
.PARAMETER CompanyFax
    追加するアドレスの会社FAX番号
.PARAMETER Link
    追加するアドレスのURL
.PARAMETER Post
    追加するアドレスの役職名
.PARAMETER Phone
    追加するアドレスの個人電話番号
.PARAMETER Email
    追加するアドレスのメールアドレス
.PARAMETER Description
    追加するアドレスのメモ
.PARAMETER URL
    ガルーンのURL
    必ずトップページのURLを指定してください
    例: http://grnserver/cgi-bin/cgi/grn.cgi
.PARAMETER Credential
    ガルーンに接続するための資格情報
.PARAMETER PassThru
    デフォルトではこの関数は処理後に何も出力しません
    -PassThruを使用すると追加されたアドレス情報を出力します
.EXAMPLE
    Add-GrnAddressBookMember -BookName '営業本部アドレス帳' -DisplayName '吉田翔平' -Email 'yoshida@example.com' -URL $URL -Credential $Credential -PassThru
    Example 1: 営業本部アドレス帳にアドレス「吉田翔平」を追加します
#>
function Add-GrnAddressBookMember {
    [CmdletBinding(DefaultParameterSetName = 'name')]
    Param
    (
        # アドレス帳の名前
        [Parameter(Mandatory = $true, ParameterSetName = 'name')]
        [string]$BookName,

        # アドレス帳ID
        [Parameter(Mandatory = $true, ParameterSetName = 'id')]
        [ValidateNotNullOrEmpty()]
        [int]$BookId,

        # アドレス表示名
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DisplayName,

        # アドレス個人名（姓）
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$FamilyName = [NullString]::Value,

        # アドレス個人名（名）
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$GivenName = [NullString]::Value,

        # アドレス個人名（よみ）（姓）
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$FamilyNameKana = [NullString]::Value,

        # アドレス個人名（よみ）（名）
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$GivenNameKana = [NullString]::Value,

        # アドレス会社名
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$CompanyName = [NullString]::Value,

        # アドレス会社名（よみ）
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$CompanyReading = [NullString]::Value,

        # アドレス部課名
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Section = [NullString]::Value,

        # アドレス郵便番号
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$ZipCode = [NullString]::Value,

        # アドレス住所
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Address = [NullString]::Value,

        # アドレス地図
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [uri]$Map = $null,

        # アドレス会社電話番号
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$CompanyPhone = [NullString]::Value,

        # アドレス会社FAX番号
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$CompanyFax = [NullString]::Value,

        # アドレスURL
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [uri]$Link = $null,

        # アドレス役職名
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Post = [NullString]::Value,

        # アドレス個人電話番号
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Phone = [NullString]::Value,

        # アドレスE-mail
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Email = [NullString]::Value,

        # アドレスメモ
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Description = [NullString]::Value,

        # ガルーンのURL
        [Parameter(Mandatory = $true)]
        [string]$URL,

        # ガルーン管理者の資格情報
        [Parameter(Mandatory = $true)]
        [pscredential]$Credential,

        [Parameter()]
        [switch]$PassThru
    )

    Begin {
        $class = New-Object GaroonAddress @($URL, $Credential) -ErrorAction Stop
        if ($PSCmdlet.ParameterSetName -eq 'name') {
            $targetBook = Get-GrnAddressBook -BookName $BookName -GetMemberInAddressBook $false -URL $URL -Credential $Credential
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'id') {
            $targetBook = Get-GrnAddressBook -BookId $BookId -GetMemberInAddressBook $false -URL $URL -Credential $Credential
        }
    }

    Process {
        if ($null -eq $targetBook) {
            Write-Error '指定されたアドレス帳が見つかりません'
            return
        }

        if (0 -eq $targetBook.BookId) {
            $targetBook.BookId = -1  #個人アドレス帳
        }

        $CardInfo = [CardInfo]::new(@{
                CardId         = 0
                Version        = 'dummy'
                DisplayName    = $DisplayName
                FamilyName     = $FamilyName
                GivenName      = $GivenName
                FamilyNameKana = $FamilyNameKana
                GivenNameKana  = $GivenNameKana
                CompanyName    = $CompanyName
                CompanyReading = $CompanyReading
                Section        = $Section
                ZipCode        = $ZipCode
                Address        = $Address
                Map            = $Map
                CompanyPhone   = $CompanyPhone
                CompanyFax     = $CompanyFax
                URL            = $Link
                Post           = $Post
                Phone          = $Phone
                Email          = $Email
                Description    = $Description
            })

        try {
            $result = $class.AddCards($CardInfo, $targetBook.BookId)
            if ($PassThru) {
                $result
            }
        }
        catch {
            Write-Error -Exception $_.Exception
        }
    }

    End {}
}
