<#
.SYNOPSIS
    ガルーンのアドレス帳に登録されたアドレスを編集します
.DESCRIPTION
    ガルーンのアドレス帳に登録されたアドレスを編集します
.PARAMETER BookName
    編集したいアドレスが含まれるアドレス帳の名前を指定します
    BookIdパラメータと同時に使用することはできません
    注意：未指定のパラメータの情報は全て削除されます
    既存のアドレス情報を維持したまま一部の情報のみ変更する場合はEXAMPLEを参考にしてください
.PARAMETER BookId
    編集したいアドレスが含まれるアドレス帳のIDを指定します
    BookNameパラメータと同時に使用することはできません
    「個人アドレス帳」を指定する場合は、BookIdに「-1」を指定してください
.PARAMETER CardId
    編集対象のアドレスIDを指定します
.PARAMETER DisplayName
    編集するアドレスの表示名
.PARAMETER FamilyName
    編集するアドレスの姓
.PARAMETER GivenName
    編集するアドレスの名
.PARAMETER FamilyNameKana
    編集するアドレスの姓（よみ）
.PARAMETER GivenNameKana
    編集するアドレスの名（よみ）
.PARAMETER CompanyName
    編集するアドレスの会社名
.PARAMETER CompanyReading
    編集するアドレスの会社名（よみ）
.PARAMETER Section
    編集するアドレスの部署名
.PARAMETER ZipCode
    編集するアドレスの郵便番号
.PARAMETER Address
    編集するアドレスの住所
.PARAMETER Map
    編集するアドレスの地図URL
.PARAMETER CompanyPhone
    編集するアドレスの会社電話番号
.PARAMETER CompanyFax
    編集するアドレスの会社FAX番号
.PARAMETER Link
    編集するアドレスのURL
.PARAMETER Post
    編集するアドレスの役職名
.PARAMETER Phone
    編集するアドレスの個人電話番号
.PARAMETER Email
    編集するアドレスのメールアドレス
.PARAMETER Description
    編集するアドレスのメモ
.PARAMETER URL
    ガルーンのURL
    必ずトップページのURLを指定してください
    例: http://grnserver/cgi-bin/cgi/grn.cgi
.PARAMETER Credential
    ガルーンに接続するための資格情報
.PARAMETER Force
    変更対象のアドレスが存在しない場合は新規に登録します
.PARAMETER PassThru
    デフォルトではこの関数は処理後に何も出力しません
    -PassThruを使用すると変更されたアドレス情報を出力します
.EXAMPLE
    Set-GrnAddressBookMember -BookName '営業本部' -CardId 3 -DisplayName '吉田翔平' -Email 'yoshida@example.com' -URL $URL -Credential $Credential -PassThru
    Example1:
      「営業本部」アドレス帳に含まれるアドレスID=3のアドレスを「吉田翔平」に変更します
      表示名とメールアドレス以外の既存情報はすべてクリアされます
.EXAMPLE
    $Member = Get-GrnAddressBook -BookName '営業本部' -URL $URL -Credential $Cred | Select-Object -ExpandProperty Member
    $Takahashi = $Member | ? {$_.FamilyName -eq '高橋'}
    $Takahashi.Email = 'newaddress@example.com'
    $Takahashi | Set-GrnAddressBookMember -BookName '営業本部' -URL $URL -Credential $Cred
    Example2:
      「営業本部」アドレス帳に含まれる姓が「高橋」のメールアドレスを「newaddress@example.com」に変更します
      メールアドレス以外の既存情報は変更されません
    )
#>
function Set-GrnAddressBookMember {
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

        # アドレスID
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [int]$CardId,

        # アドレス表示名
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$DisplayName = [NullString]::Value,

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
        [switch]$Force,

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

        if ($targetBook.CardId -notcontains $CardId) {
            if (-not $Force) {
                Write-Error ('指定されたアドレス帳に CardId={0} のアドレスが存在しません' -f $CardId)
                return
            }
            else {
                if ([string]::IsNullOrEmpty($DisplayName)) {
                    Write-Error 'DisplayNameを指定する必要があります'
                    return
                }
                else {
                    $AddParam = $PSBoundParameters
                    $AddParam.Remove('CardId') >$null
                    $AddParam.Remove('Force') >$null
                    Add-GrnAddressBookMember @PSBoundParameters
                    return
                }
            }
        }

        if (0 -eq $targetBook.BookId) {
            $targetBook.BookId = -1  #個人アドレス帳
        }

        $CardInfo = [CardInfo]::new(@{
                CardId         = $CardId
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
            $result = $class.ModifyCard($CardInfo, $CardId, $targetBook.BookId)
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
