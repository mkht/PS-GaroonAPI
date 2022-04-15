using namespace System.Xml
using namespace System.Security

class CardInfo {
    [string]$BookId = [NullString]::Value  # ブックID
    [string]$CardId = [NullString]::Value  # カードID
    [string]$Version = [NullString]::Value  # バージョン
    [pscustomobject]$Creator = $null   # 作成者
    [pscustomobject]$Modifier = $null   # 更新者
    [string]$DisplayName = [NullString]::Value    # 表示名
    [string]$FamilyName = [NullString]::Value    # 個人名（姓）
    [string]$GivenName = [NullString]::Value    # 個人名（名）
    [string]$FamilyNameKana = [NullString]::Value    # 個人名（よみ）（姓）
    [string]$GivenNameKana = [NullString]::Value    # 個人名（よみ）（名）
    [string]$CompanyName = [NullString]::Value    # 会社名
    [string]$CompanyReading = [NullString]::Value    # 会社名（よみ）
    [string]$Section = [NullString]::Value    # 部課名
    [string]$ZipCode = [NullString]::Value    # 郵便番号
    [string]$Address = [NullString]::Value    # 住所
    [Uri]$Map = $null   # 地図
    [string]$CompanyPhone = [NullString]::Value    # 会社電話番号
    [string]$CompanyFax = [NullString]::Value    # 会社FAX番号
    [Uri]$URL = $null   # URL
    [string]$Post = [NullString]::Value    # 役職名
    [string]$Phone = [NullString]::Value    # 個人電話番号
    [string]$Email = [NullString]::Value    # E-mail
    [string]$Description = [NullString]::Value    # メモ

    <# ---- コンストラクタ ---- #>
    CardInfo() {}

    CardInfo([Object]$Info) {
        if (-not [string]::IsNullOrWhiteSpace($Info.BookId)) { $this.BookId = ([string]$Info.BookId).Trim() }
        if (-not [string]::IsNullOrWhiteSpace($Info.CardId)) { $this.CardId = ([string]$Info.CardId).Trim() }
        if ($null -ne $Info.Version) { $this.Version = $Info.Version }
        if ($null -ne $Info.Creator) {
            $this.Creator = [pscustomobject]@{
                user_id = ([string]$Info.Creator.user_id).Trim()
                name    = [string]$Info.Creator.name
                date    = [datetime]$Info.Creator.date
            }
        }
        if ($null -ne $Info.Modifier) {
            $this.Modifier = [pscustomobject]@{
                user_id = ([string]$Info.Modifier.user_id).Trim()
                name    = [string]$Info.Modifier.name
                date    = [datetime]$Info.Modifier.date
            }
        }
        if ($null -ne $Info.DisplayName) { $this.DisplayName = $Info.DisplayName }
        if ($null -ne $Info.FamilyName) { $this.FamilyName = $Info.FamilyName }
        if ($null -ne $Info.GivenName) { $this.GivenName = $Info.GivenName }
        if ($null -ne $Info.FamilyNameKana) { $this.FamilyNameKana = $Info.FamilyNameKana }
        if ($null -ne $Info.GivenNameKana) { $this.GivenNameKana = $Info.GivenNameKana }
        if ($null -ne $Info.CompanyName) { $this.CompanyName = $Info.CompanyName }
        if ($null -ne $Info.CompanyReading) { $this.CompanyReading = $Info.CompanyReading }
        if ($null -ne $Info.Section) { $this.Section = $Info.Section }
        if ($null -ne $Info.ZipCode) { $this.ZipCode = $Info.ZipCode }
        if ($null -ne $Info.Address) { $this.Address = $Info.Address }
        if ($Info.Map -as [uri]) { $this.Map = [uri]$Info.Map }
        if ($null -ne $Info.CompanyPhone) { $this.CompanyPhone = $Info.CompanyPhone }
        if ($null -ne $Info.CompanyFax) { $this.CompanyFax = $Info.CompanyFax }
        if ($Info.URL -as [uri]) { $this.URL = [uri]$Info.URL }
        if ($null -ne $Info.Post) { $this.Post = $Info.Post }
        if ($null -ne $Info.Phone) { $this.Phone = $Info.Phone }
        if ($null -ne $Info.Email) { $this.Email = $Info.Email }
        if ($null -ne $Info.Description) { $this.Description = $Info.Description }
    }

    [string]GetCardInfoString() {
        $attr = @()
        $elem = @()
        if ($null -ne $this.BookId) {
            $attr += ('book_id="{0}"' -f [SecurityElement]::Escape($this.BookId))
        }
        if ($null -ne $this.CardId) {
            $attr += ('id="{0}"' -f [SecurityElement]::Escape($this.CardId))
        }
        if ($null -ne $this.Version) {
            $attr += ('version="{0}"' -f [SecurityElement]::Escape($this.Version))
        }
        if ($null -ne $this.DisplayName) {
            $elem += ('<subject>{0}</subject>' -f [SecurityElement]::Escape($this.DisplayName))
        }
        if ($null -ne $this.FamilyName) {
            $elem += ('<personal_name><part>{0}</part><part>{1}</part></personal_name>' -f [SecurityElement]::Escape($this.FamilyName), [SecurityElement]::Escape($this.GivenName))
        }
        if ($null -ne $this.FamilyNameKana) {
            $elem += ('<personal_reading><part>{0}</part><part>{1}</part></personal_reading>' -f [SecurityElement]::Escape($this.FamilyNameKana), [SecurityElement]::Escape($this.GivenNameKana))
        }
        if ($null -ne $this.CompanyName) {
            $elem += ('<company_name>{0}</company_name>' -f [SecurityElement]::Escape($this.CompanyName))
        }
        if ($null -ne $this.CompanyReading) {
            $elem += ('<company_reading>{0}</company_reading>' -f [SecurityElement]::Escape($this.CompanyReading))
        }
        if ($null -ne $this.Section) {
            $elem += ('<section>{0}</section>' -f [SecurityElement]::Escape($this.Section))
        }
        if ($null -ne $this.ZipCode) {
            $elem += ('<zip_code>{0}</zip_code>' -f [SecurityElement]::Escape($this.ZipCode))
        }
        if ($null -ne $this.Address) {
            $elem += ('<physical_address>{0}</physical_address>' -f [SecurityElement]::Escape($this.Address))
        }
        if ($null -ne $this.Map) {
            $elem += ('<map>{0}</map>' -f [SecurityElement]::Escape($this.Map.OriginalString))
        }
        if ($null -ne $this.CompanyPhone) {
            $elem += ('<company_tel>{0}</company_tel>' -f [SecurityElement]::Escape($this.CompanyPhone))
        }
        if ($null -ne $this.CompanyFax) {
            $elem += ('<company_fax>{0}</company_fax>' -f [SecurityElement]::Escape($this.CompanyFax))
        }
        if ($null -ne $this.URL) {
            $elem += ('<url>{0}</url>' -f [SecurityElement]::Escape($this.URL.OriginalString))
        }
        if ($null -ne $this.Post) {
            $elem += ('<post>{0}</post>' -f [SecurityElement]::Escape($this.Post))
        }
        if ($null -ne $this.Phone) {
            $elem += ('<personal_tel>{0}</personal_tel>' -f [SecurityElement]::Escape($this.Phone))
        }
        if ($null -ne $this.Email) {
            $elem += ('<email>{0}</email>' -f [SecurityElement]::Escape($this.Email))
        }
        if ($null -ne $this.Description) {
            $elem += ('<description>{0}</description>' -f [SecurityElement]::Escape($this.Description))
        }

        return [string]('<card {0}>{1}</card>' -f ($attr -join ' '), ($elem -join ''))
    }
}

class BookInfo {
    [string]$BookId = [NullString]::Value  # ブックID
    [string]$Key = [NullString]::Value  # ブックのコード
    [string]$Name = [NullString]::Value  # 	ブック名
    [string]$Version = [NullString]::Value  # バージョン
    [string]$Type = [NullString]::Value  # 	ブックタイプ
    [string[]]$CardId = @()  # カードID
    [pscustomobject]$Form = $null   # ブック項目の一覧

    <# ---- コンストラクタ ---- #>
    BookInfo() {}

    BookInfo([Object]$Info) {
        if (-not [string]::IsNullOrWhiteSpace($Info.BookId)) { $this.BookId = ([string]$Info.BookId).Trim() }
        if ($null -ne $Info.Key) { $this.Key = $Info.Key }
        if ($null -ne $Info.Name) { $this.Name = $Info.Name }
        if ($null -ne $Info.Version) { $this.Version = $Info.Version }
        if ($null -ne $Info.Type) { $this.Type = $Info.Type }
        if ($null -ne $Info.CardId) { $this.CardId = $Info.CardId }
        if ($null -ne $Info.Form) { $this.Form = [pscustomobject]$Info.Form }
    }
}


# アドレス帳API群実行用クラス
# https://developer.cybozu.io/hc/ja/articles/202251654
Class GaroonAddress : GaroonClass {
    [string] $ApiSuffix = '/cbpapi/address/api'
    GaroonAddress() : base() {}
    GaroonAddress([string]$URL) : base($URL) {}
    GaroonAddress([string]$URL, [PSCredential] $Credential) : base($URL, $Credential) {}

    #共有アドレス帳を取得する
    [BookInfo[]]GetSharedBooksById([int[]]$BookId) {
        $Action = 'AddressGetSharedBooksById'
        [string[]]$body = $BookId | ForEach-Object { '<book_id>{0}</book_id>' -f $_ }
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.AddressGetSharedBooksByIdResponse.returns.book |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [BookInfo]::new( @{
                    BookId  = $_.book_id
                    Key     = $_.key
                    Name    = $_.name
                    Version = $_.version
                    Type    = $_.type
                    CardId  = $(if ($null -ne $_.cards.card.id) { $_.cards.card.id })
                    Form    = [pscustomobject]$_.form
                }
            )
        }
    }

    #個人アドレス帳を取得する
    [BookInfo[]]GetPersonalBooksById() {
        return $this.GetPersonalBooksById(1)  #個人アドレス帳を複数持つケースがあるのか不明
    }

    [BookInfo[]]GetPersonalBooksById([int[]]$BookId) {
        $Action = 'AddressGetPersonalBooksById'
        [string[]]$body = $BookId | ForEach-Object { '<book_id>{0}</book_id>' -f $_ }
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.AddressGetPersonalBooksByIdResponse.returns.book |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [BookInfo]::new( @{
                    BookId  = $_.book_id
                    Key     = $_.key
                    Name    = $_.name
                    Version = $_.version
                    Type    = $_.type
                    CardId  = $(if ($null -ne $_.cards.card.id) { $_.cards.card.id })
                    Form    = [pscustomobject]$_.form
                }
            )
        }
    }

    #共有アドレス帳のアドレスを取得する
    [CardInfo[]]GetSharedCardsById([int[]]$CardId, [int]$BookId) {
        $Action = 'AddressGetSharedCardsById'
        [string[]]$body = $CardId | ForEach-Object { '<card_id>{0}</card_id>' -f $_ }
        $ParamBody = ('<parameters book_id="{1}">{0}</parameters>' -f ($body -join ''), $BookId)
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.AddressGetSharedCardsByIdResponse.returns.card |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [CardInfo]::new( @{
                    BookId         = $_.book_id
                    CardId         = $_.id
                    Version        = $_.version
                    Creator        = [pscustomobject]$_.creator
                    Modifier       = [pscustomobject]$_.modifier
                    DisplayName    = $_.subject
                    FamilyName     = $(if ($_.personal_name.part -is [array]) { $_.personal_name.part[0] })
                    GivenName      = $(if ($_.personal_name.part -is [array]) { $_.personal_name.part[1] })
                    FamilyNameKana = $(if ($_.personal_reading.part -is [array]) { $_.personal_reading.part[0] })
                    GivenNameKana  = $(if ($_.personal_reading.part -is [array]) { $_.personal_reading.part[1] })
                    CompanyName    = $_.company_name
                    CompanyReading = $_.company_reading
                    Section        = $_.section
                    ZipCode        = $_.zip_code
                    Address        = $_.physical_address
                    Map            = [uri]$_.map
                    CompanyPhone   = $_.company_tel
                    CompanyFax     = $_.company_fax
                    URL            = [uri]$_.url
                    Post           = $_.post
                    Phone          = $_.personal_tel
                    Email          = $_.email
                    Description    = $_.description
                }
            )
        }
    }

    #個人アドレス帳のアドレスを取得する
    [CardInfo[]]GetPersonalCardsById([int[]]$CardId) {
        $Action = 'AddressGetPersonalCardsById'
        [string[]]$body = $CardId | ForEach-Object { '<card_id>{0}</card_id>' -f $_ }
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.AddressGetPersonalCardsByIdResponse.returns.card |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [CardInfo]::new( @{
                    BookId         = $_.book_id
                    CardId         = $_.id
                    Version        = $_.version
                    Creator        = [pscustomobject]$_.creator
                    Modifier       = [pscustomobject]$_.modifier
                    DisplayName    = $_.subject
                    FamilyName     = $(if ($_.personal_name.part -is [array]) { $_.personal_name.part[0] })
                    GivenName      = $(if ($_.personal_name.part -is [array]) { $_.personal_name.part[1] })
                    FamilyNameKana = $(if ($_.personal_reading.part -is [array]) { $_.personal_reading.part[0] })
                    GivenNameKana  = $(if ($_.personal_reading.part -is [array]) { $_.personal_reading.part[1] })
                    CompanyName    = $_.company_name
                    CompanyReading = $_.company_reading
                    Section        = $_.section
                    ZipCode        = $_.zip_code
                    Address        = $_.physical_address
                    Map            = [uri]$_.map
                    CompanyPhone   = $_.company_tel
                    CompanyFax     = $_.company_fax
                    URL            = [uri]$_.url
                    Post           = $_.post
                    Phone          = $_.personal_tel
                    Email          = $_.email
                    Description    = $_.description
                }
            )
        }
    }

    #アドレスを検索する
    [CardInfo[]]SearchCards([string]$SearchText, [int]$BookId) {
        return $this.SearchCards($SearchText, $BookId, $true)
    }

    [CardInfo[]]SearchCards([string]$SearchText, [int]$BookId, [bool]$IgnoreCase) {
        $Action = 'AddressSearchCards'
        $ParamBody = ('<parameters book_id="{0}" case_sensitive="{1}" text="{2}"></parameters>' -f $BookId, (-not $IgnoreCase).ToString().ToLower(), [SecurityElement]::Escape($SearchText))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.AddressSearchCardsResponse.returns.card |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [CardInfo]::new( @{
                    BookId         = $_.book_id
                    CardId         = $_.id
                    Version        = $_.version
                    Creator        = [pscustomobject]$_.creator
                    Modifier       = [pscustomobject]$_.modifier
                    DisplayName    = $_.subject
                    FamilyName     = $(if ($_.personal_name.part -is [array]) { $_.personal_name.part[0] })
                    GivenName      = $(if ($_.personal_name.part -is [array]) { $_.personal_name.part[1] })
                    FamilyNameKana = $(if ($_.personal_reading.part -is [array]) { $_.personal_reading.part[0] })
                    GivenNameKana  = $(if ($_.personal_reading.part -is [array]) { $_.personal_reading.part[1] })
                    CompanyName    = $_.company_name
                    CompanyReading = $_.company_reading
                    Section        = $_.section
                    ZipCode        = $_.zip_code
                    Address        = $_.physical_address
                    Map            = [uri]$_.map
                    CompanyPhone   = $_.company_tel
                    CompanyFax     = $_.company_fax
                    URL            = [uri]$_.url
                    Post           = $_.post
                    Phone          = $_.personal_tel
                    Email          = $_.email
                    Description    = $_.description
                }
            )
        }
    }

    #アドレスを登録する
    [CardInfo[]]AddCards([Object[]]$Card, [int]$BookId) {
        $Card | ForEach-Object {
            if ($_.GetType().FullName -ne 'CardInfo') { return $null } # Type check
            $_.BookId = $BookId
        }
        $Action = 'AddressAddCards'
        [string[]]$body = $Card | ForEach-Object { '<add_card>{0}</add_card>' -f $_.GetCardInfoString() }
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.AddressAddCardsResponse.returns.card |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [CardInfo]::new( @{
                    BookId         = $_.book_id
                    CardId         = $_.id
                    Version        = $_.version
                    Creator        = [pscustomobject]$_.creator
                    Modifier       = [pscustomobject]$_.modifier
                    DisplayName    = $_.subject
                    FamilyName     = $(if ($_.personal_name.part -is [array]) { $_.personal_name.part[0] })
                    GivenName      = $(if ($_.personal_name.part -is [array]) { $_.personal_name.part[1] })
                    FamilyNameKana = $(if ($_.personal_reading.part -is [array]) { $_.personal_reading.part[0] })
                    GivenNameKana  = $(if ($_.personal_reading.part -is [array]) { $_.personal_reading.part[1] })
                    CompanyName    = $_.company_name
                    CompanyReading = $_.company_reading
                    Section        = $_.section
                    ZipCode        = $_.zip_code
                    Address        = $_.physical_address
                    Map            = [uri]$_.map
                    CompanyPhone   = $_.company_tel
                    CompanyFax     = $_.company_fax
                    URL            = [uri]$_.url
                    Post           = $_.post
                    Phone          = $_.personal_tel
                    Email          = $_.email
                    Description    = $_.description
                }
            )
        }
    }

    #アドレスを更新する
    [CardInfo]ModifyCard([Object]$Card, [int]$CardId, [int]$BookId) {
        # Type check
        if ($Card -is [array]) {
            if ($Card.Count -ne 1) {
                throw [ArgumentException]::new('Card must not be an array of objects.')
                return $null
            }
            else {
                $Card = $Card[0]
            }
        }
        if ($Card.GetType().FullName -ne 'CardInfo') { return $null }

        $Card.CardId = $CardId
        $Card.BookId = $BookId
        $Action = 'AddressModifyCards'
        [string[]]$body = $Card | ForEach-Object { '<modify_card>{0}</modify_card>' -f $_.GetCardInfoString() }
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.AddressModifyCardsResponse.returns.card |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [CardInfo]::new( @{
                    BookId         = $_.book_id
                    CardId         = $_.id
                    Version        = $_.version
                    Creator        = [pscustomobject]$_.creator
                    Modifier       = [pscustomobject]$_.modifier
                    DisplayName    = $_.subject
                    FamilyName     = $(if ($_.personal_name.part -is [array]) { $_.personal_name.part[0] })
                    GivenName      = $(if ($_.personal_name.part -is [array]) { $_.personal_name.part[1] })
                    FamilyNameKana = $(if ($_.personal_reading.part -is [array]) { $_.personal_reading.part[0] })
                    GivenNameKana  = $(if ($_.personal_reading.part -is [array]) { $_.personal_reading.part[1] })
                    CompanyName    = $_.company_name
                    CompanyReading = $_.company_reading
                    Section        = $_.section
                    ZipCode        = $_.zip_code
                    Address        = $_.physical_address
                    Map            = [uri]$_.map
                    CompanyPhone   = $_.company_tel
                    CompanyFax     = $_.company_fax
                    URL            = [uri]$_.url
                    Post           = $_.post
                    Phone          = $_.personal_tel
                    Email          = $_.email
                    Description    = $_.description
                }
            )
        }
    }

    #共有アドレス帳のアドレスを削除する
    [void]RemoveSharedCards([int[]]$CardId, [int]$BookId) {
        $Action = 'AddressRemoveSharedCards'
        [string[]]$body = $CardId | ForEach-Object { '<card_id>{0}</card_id>' -f $_ }
        $ParamBody = ('<parameters book_id="{1}">{0}</parameters>' -f ($body -join ''), $BookId)
        $null = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
    }

    #個人アドレス帳のアドレスを削除する
    [void]RemovePersonalCards([int[]]$CardId) {
        $Action = 'AddressRemovePersonalCards'
        [string[]]$body = $CardId | ForEach-Object { '<card_id>{0}</card_id>' -f $_ }
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $null = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
    }

    #閲覧可能なアドレス帳を確認する
    [int[]]GetReadAllowBooks() {
        $Action = 'AddressGetReadAllowBooks'
        $ParamBody = '<parameters></parameters>'
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.AddressGetReadAllowBooksResponse.returns.book_id
    }

    #編集可能なアドレス帳を確認する
    [int[]]GetModifyAllowBooks() {
        $Action = 'AddressGetModifyAllowBooks'
        $ParamBody = '<parameters></parameters>'
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.AddressGetModifyAllowBooksResponse.returns.book_id
    }
}
