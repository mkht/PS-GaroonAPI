<#
.Synopsis
   ガルーンのアドレス帳を取得します
.DESCRIPTION
   ブック名もしくはブックIDをキーにガルーンのアドレス帳を取得します。
   ブック名に何も指定しない場合、閲覧可能なすべてのアドレス帳を取得します。
.EXAMPLE
    Set-GrnAddressBookMember -BookName '営業本部' -URL 'http://grnserver/grn.cgi' -Credential (Get-Credential)

    BookId  : 3
    Key     : Sales Headquarters
    Name    : 営業本部
    Version : 1191475956
    Type    : cyde
    CardId  : {3, 6, 7, 4}
    Member  : {加藤 美咲, 鈴木 卓也, 音無 結城, 高橋 健太}
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

        # アドレス帳ID
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [int]$CardId,

        # アドレス表示名
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DisplayName,

        # アドレス個人名（姓）
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$FamilyName,

        # アドレス個人名（名）
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$GivenName,

        # アドレス個人名（よみ）（姓）
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$FamilyNameKana,

        # アドレス個人名（よみ）（名）
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$GivenNameKana,

        # アドレス会社名
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$CompanyName,

        # アドレス会社名（よみ）
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$CompanyReading,

        # アドレス部課名
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Section,

        # アドレス郵便番号
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$ZipCode,

        # アドレス住所
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Address,

        # アドレス地図
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [uri]$Map,

        # アドレス会社電話番号
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$CompanyPhone,

        # アドレス会社FAX番号
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$CompanyFax,

        # アドレスURL
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [uri]$Link,

        # アドレス役職名
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Post,

        # アドレス個人電話番号
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Phone,

        # アドレスE-mail
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Email,

        # アドレスメモ
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Description,

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
                $AddParam = $PSBoundParameters
                $AddParam.Remove('CardId')
                $AddParam.Remove('Force')
                Add-GrnAddressBookMember @PSBoundParameters
                return
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
