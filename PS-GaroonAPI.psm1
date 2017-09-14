class GaroonClass {
    [string]$RequestBase = @'
<?xml version="1.0" encoding="UTF-8"?>
<soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
  <soap:Header>
    <Action>
      {0}
    </Action>
    <Security>
      <UsernameToken>
        <Username>{2}</Username>
        <Password>{3}</Password>
      </UsernameToken>
    </Security>
    <Timestamp>
      <Created>{4}</Created>
      <Expires>{5}</Expires>
    </Timestamp>
    <Locale>jp</Locale>
  </soap:Header>
  <soap:Body>
  <{0}>
    {1}
  </{0}>
  </soap:Body>
</soap:Envelope>
'@
    [System.Management.Automation.PSCredential] $Credential # ガルーンログインアカウント情報
    [ValidatePattern('^https?://(.+(grn|index)(\.exe|\.cgi|\.csp)\??|.+\.cybozu\.com/g/)$')]
    [string] $GrnURL    # ガルーンのURL
    [ValidateSet(".exe", ".csp", ".cgi", "")]
    [string] $Ext   # [未使用]拡張子
    [string] $ApiSuffix
    [string] $RequestURI    # 実際にリクエストするURL
    <# ---- コンストラクタ ---- #>
    GaroonClass() {}
    GaroonClass([string]$URL) {
        $this.GrnURL = $URL
        if ($tURL = $([regex]::Match($URL, ".+cybozu.com/g"))[0].Value) {
            $this.RequestURI = $tURL + $this.ApiSuffix
        }
        else {
            $this.RequestURI = $this.GrnURL + $this.ApiSuffix
        }
    }
    GaroonClass([string]$URL, [PSCredential] $Credential) {
        $this.GrnURL = $URL
        $this.Credential = $Credential
        if ($tURL = $([regex]::Match($URL, ".+cybozu.com/g"))[0].Value) {
            $this.RequestURI = $tURL + $this.ApiSuffix
        }
        else {
            $this.RequestURI = $this.GrnURL + $this.ApiSuffix
        }
    }
    static [xml]InvokeSOAPRequest ([Xml]$SOAPRequest, [String] $URL) {
        Write-Verbose "Sending SOAP Request To Server: $URL"
        $soapWebRequest = [System.Net.WebRequest]::Create($URL)
        $soapWebRequest.ContentType = 'text/xml;charset="utf-8"'
        $soapWebRequest.Accept = "text/xml"
        $soapWebRequest.Method = "POST"
        Write-Verbose "Initiating Send."
        $requestStream = $soapWebRequest.GetRequestStream()
        $SOAPRequest.Save($requestStream)
        $requestStream.Close()
        Write-Verbose "Send Complete, Waiting For Response."
        $resp = $soapWebRequest.GetResponse()
        $responseStream = $resp.GetResponseStream()
        $soapReader = [System.IO.StreamReader]($responseStream)
        $ReturnXml = [Xml] $soapReader.ReadToEnd()
        $responseStream.Close()
        Write-Verbose "Response Received."
        return $ReturnXml
    }
    [string]CreateRequestXml([string]$Action, [string]$ParamterBody, [DateTime]$CreateTime) {
        $CreateUtcTime = [System.TimeZoneInfo]::ConvertTimeToUtc($CreateTime)   # GaroonAPIのTimestampはUTCのみ
        $ExpireUtcTime = $CreateUtcTime.AddHours(1) #期限は適当に1時間
        if ($this.Credential) {
            return ($this.RequestBase -f $Action, $ParamterBody, $this.Credential.UserName, $this.Credential.GetNetworkCredential().Password, $CreateUtcTime.ToString('yyyy-MM-ddTHH:mm:ssK'), $ExpireUtcTime.ToString('yyyy-MM-ddTHH:mm:ssK'))
        }
        else {
            return ($this.RequestBase -f $Action, $ParamterBody, '', '', $CreateUtcTime.ToString('yyyy-MM-ddTHH:mm:ssK'), $ExpireUtcTime.ToString('yyyy-MM-ddTHH:mm:ssK'))
        }
    }
    [xml]Request([string]$RequestXml) {
        return $this.TestReponseXml([GaroonClass]::InvokeSOAPRequest($RequestXml, $this.RequestURI))
    }
    [Object[]]InvokeAnyApi($Action, $RequestBody) {
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $RequestBody, (Get-Date)))
        return $ReponseXml.Envelope.Body
    }
    [xml]TestReponseXml([xml]$ReponseXml) {
        if (-not $ReponseXml) {
            $msg = 'No Response returned.'
            throw $msg    # throwとWrite-Errorどっちがいいんだろう？
        }
        if ($Fault = $ReponseXml.Envelope.Body.Fault) {
            $msg = ("[ERROR][{0}] {1}" -f $Fault.Detail.Code, $Fault.Reason.Text)
            throw $msg
        }
        return $ReponseXml
    }
}
class MailAccountInfo {
    [string]$AccountId   # アカウントID
    [string]$UserId  # ユーザーID
    [string]$UserAccountCode    # アカウントコード
    [string]$UserAccountName    # 	アカウント名
    <# ---- コンストラクタ ---- #>
    MailAccountInfo() {}
    MailAccountInfo([Object]$Info) {
        $this.AccountId = $Info.AccountId
        $this.UserId = $Info.UserId
        $this.UserAccountCode = $Info.UserAccountCode
        $this.UserAccountName = $Info.UserAccountName
    }
    [string]GetAccountInfoString() {
        $attr = @()
        $attr += ('account_id="{0}"' -f [string]$this.AccountId)
        $attr += ('user_id="{0}"' -f $this.UserId)
        $attr += ('user_acount_code="{0}"' -f $this.UserAccountCode)  # スペルミスはGaronnAPIの仕様
        if ($this.UserAccountName) {
            $attr += ('user_account_name="{0}"' -f $this.UserAccountName)
        }
        return [string]('<account_info {0}></account_info>' -f ($attr -join ' '))
    }
}
class MailSetting {
    [string]$MailServerId   # メールサーバーID
    [string]$Email  # メールアドレス
    [string]$AccountName  # アカウント名
    [string]$Password    # パスワード
    [bool]$LeaveServerMail = $false   # メールサーバーにメールを残す
    [bool]$Deactivate = $false    # アカウント停止
    <# ---- コンストラクタ ---- #>
    MailSetting() {}
    MailSetting([Object]$Info) {
        $this.MailServerId = $Info.MailServerId
        $this.Email = $Info.Email
        $this.AccountName = $Info.AccountName
        $this.Password = $Info.Password
        if ($null -ne $Info.LeaveServerMail) {
            $this.LeaveServerMail = $Info.LeaveServerMail
        }
        if ($null -ne $Info.Deactivate) {
            $this.Deactivate = $Info.Deactivate
        }
    }
    [string]GetMailSettingString() {
        $attr = @()
        $attr += ('mail_server_id="{0}"' -f [string]$this.MailServerId)
        $attr += ('email="{0}"' -f $this.Email)
        $attr += ('acount_name="{0}"' -f $this.AccountName)  # スペルミスはGaronnAPIの仕様
        if ($this.Password) {$attr += ('password="{0}"' -f $this.Password)}
        $attr += ('leave_server_mail="{0}"' -f $this.LeaveServerMail.ToString().ToLower())
        $attr += ('deactivate_user_account="{0}"' -f $this.Deactivate.ToString().ToLower())
        return [string]('<mail_setting {0}></mail_setting>' -f ($attr -join ' '))
    }
}
class MailUserAccount {
    [MailAccountInfo]$MailAccountInfo
    [MailSetting]$MailSetting
    <# ---- コンストラクタ ---- #>
    MailUserAccount() {
        $this.MailAccountInfo = [MailAccountInfo]::new()
        $this.MailSetting = [MailSetting]::new()
    }
    MailUserAccount([MailAccountInfo]$MailAccountInfo, [MailSetting]$MailSetting) {
        $this.MailAccountInfo = $MailAccountInfo
        $this.MailSetting = $MailSetting
    }
    MailUserAccount([Object]$Info) {
        $this.MailAccountInfo = [MailAccountInfo]::new($Info)
        $this.MailSetting = [MailSetting]::new($Info)
    }
    [string]GetMailUserAccountString() {
        return [string]('{0}{1}' -f $this.MailAccountInfo.GetAccountInfoString(), $this.MailSetting.GetMailSettingString())
    }
}
Class GaroonMail : GaroonClass {
    [string] $ApiSuffix = "/cbpapi/mail/api"
    GaroonMail() : base() {}
    GaroonMail([string]$URL) : base($URL) {}
    GaroonMail([string]$URL, [PSCredential] $Credential) : base($URL, $Credential) {}
    [MailUserAccount[]]CreateUserAccount([Object[]]$MailUserAccount) {
        $MailUserAccount | ForEach-Object {
            if ($_.GetType().FullName -ne "MailUserAccount") {return $null} # Type check
        }
        $Action = "MailCreateUserAccount"
        [string[]]$body = $MailUserAccount | Foreach-Object {('<mail_user_accounts xmlns="">{0}</mail_user_accounts>' -f $_.GetMailUserAccountString())}
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.MailCreateUserAccountResponse.returns.user_accounts |
            ForEach-Object {[MailUserAccount]::new( @{
                    AccountId       = $_.account_info.account_id;
                    UserId          = $_.account_info.user_id;
                    UserAccountCode = $_.account_info.user_acount_code
                    UserAccountName = $_.account_info.user_account_name
                    MailServerId    = $_.mail_setting.mail_server_id
                    Email           = $_.mail_setting.email
                    AccountName     = $_.mail_setting.acount_name
                    Password        = $_.mail_setting.password
                    LeaveServerMail = [bool]$_.mail_setting.leave_server_mail
                    Deactivate      = [bool]$_.mail_setting.deactivate_user_account
                }
            )
        }
    }
    [MailUserAccount[]]EditUserAccount([Object[]]$MailUserAccount) {
        $MailUserAccount | ForEach-Object {
            if ($_.GetType().FullName -ne "MailUserAccount") {return $null} # Type check
        }
        $Action = "MailEditUserAccount"
        [string[]]$body = $MailUserAccount | Foreach-Object {'<edit_user_accounts xmlns="">{0}</edit_user_accounts>' -f $_.GetMailUserAccountString()}
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.MailEditUserAccountResponse.returns.user_accounts |
            ForEach-Object {[MailUserAccount]::new( @{
                    AccountId       = $_.account_info.account_id;
                    UserId          = $_.account_info.user_id;
                    UserAccountCode = $_.account_info.user_acount_code
                    UserAccountName = $_.account_info.user_account_name
                    MailServerId    = $_.mail_setting.mail_server_id
                    Email           = $_.mail_setting.email
                    AccountName     = $_.mail_setting.acount_name
                    Password        = $_.mail_setting.password
                    LeaveServerMail = [bool]$_.mail_setting.leave_server_mail
                    Deactivate      = [bool]$_.mail_setting.deactivate_user_account
                }
            )
        }
    }
    [void]DeleteUserAccount([string]$AccountId, [bool]$DeleteAllEmail) {
        $Action = "MailDeleteUserAccount"
        $ParamBody = ('<parameters><delete_user_accounts xmlns="" account_id="{0}" delete_all_email="{1}"></delete_user_accounts></parameters>' -f $AccountId, $DeleteAllEmail.ToString().ToLower())
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return
    }
    [Object[]]GetAccountsById([string[]]$AccountId) {
        $Action = "MailGetAccountsById"
        [string[]]$body = $AccountId | Foreach-Object {'<account_id xmlns="">{0}</account_id> ' -f $_}
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.MailGetAccountsByIdResponse.returns.account
    }
}
Class GaroonBase : GaroonClass {
    [string] $ApiSuffix = "/cbpapi/base/api"
    GaroonBase() : base() {}
    GaroonBase([string]$URL) : base($URL) {}
    GaroonBase([string]$URL, [PSCredential] $Credential) : base($URL, $Credential) {}
    [Object[]]GetApplicationStatus() {
        $Action = "BaseGetApplicationStatus"
        $ParamBody = '<parameters></parameters>'
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.BaseGetApplicationStatusResponse.returns.application
    }
    [Object[]]GetUsersById([int[]]$UserId) {
        $Action = "BaseGetUsersById"
        [string[]]$body = $UserId | Foreach-Object {"<user_id>{0}</user_id>" -f $_}
        $ParamBody = ('<parameters xmlns="">{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.BaseGetUsersByIdResponse.returns.user
    }
    [Object[]]GetUsersByLoginName([string[]]$LoginName) {
        $Action = "BaseGetUsersByLoginName"
        [string[]]$body = $LoginName | Foreach-Object {"<login_name>{0}</login_name>" -f $_}
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.BaseGetUsersByLoginNameResponse.returns.user
    }
    [Object[]]GetOrganizationsById([int[]]$OrganizationId) {
        $Action = "BaseGetOrganizationsById"
        [string[]]$body = $OrganizationId | Foreach-Object {"<organization_id>{0}</organization_id>" -f $_}
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.BaseGetOrganizationsByIdResponse.returns.organization
    }
}
class UserInfo {
    [string]$PrimaryGroup = [NullString]::Value   # 優先する組織のID
    [int]$Position = -1  # 表示優先度
    [bool]$Invalid = $false # 使用状態(true:停止, false:使用)
    [string]$SortKey = [NullString]::Value    # よみ
    [string]$EmailAddress = [NullString]::Value   # メールアドレス
    [string]$Description = [NullString]::Value    # メモ
    [string]$Post = [NullString]::Value   # 役職
    [string]$TelNumber = [NullString]::Value  # 電話番号
    [string]$Url = [NullString]::Value    # URL
    [string]$Locale = [NullString]::Value # ロケールID
    [string]$Base = [NullString]::Value   # 拠点ID
    [int[]]$Organization    # ユーザーが所属する組織のID一覧
    <# ---- コンストラクタ ---- #>
    UserInfo() {}
    UserInfo([Object]$Info) {
        if ($null -ne $Info.PrimaryGroup) { $this.PrimaryGroup = $Info.PrimaryGroup }
        if ($null -ne $Info.Position) { $this.Position = $Info.Position }
        if ($null -ne $Info.Invalid) { $this.Invalid = $Info.Invalid }
        if ($null -ne $Info.SortKey) { $this.SortKey = $Info.SortKey }
        if ($null -ne $Info.EmailAddress) { $this.EmailAddress = $Info.EmailAddress }
        if ($null -ne $Info.Description) { $this.Description = $Info.Description }
        if ($null -ne $Info.Post) { $this.Post = $Info.Post }
        if ($null -ne $Info.TelNumber) { $this.TelNumber = $Info.TelNumber }
        if ($null -ne $Info.Url) { $this.Url = $Info.Url }
        if ($null -ne $Info.Locale) { $this.Locale = $Info.Locale }
        if ($null -ne $Info.Base) { $this.Base = $Info.Base }
        if ($null -ne $Info.Organization) { $this.Organization = $Info.Organization }
    }
    [string]GetUserInfoString() {
        $attr = @()
        if ($this.Position -as [System.UInt32]) {
            $attr += ('position="{0}"' -f [string]$this.Position)
        }
        if ($this.Invalid -is [bool]) {
            $attr += ('invalid="{0}"' -f ([string]$this.Invalid).ToLower())
        }
        if ($this.SortKey -ne $null) {
            $attr += ('sort_key="{0}"' -f $this.SortKey)
        }
        if ($this.EmailAddress -ne $null) {
            $attr += ('email_address="{0}"' -f $this.EmailAddress)
        }
        if ($this.Description -ne $null) {
            $attr += ('description="{0}"' -f $this.Description)
        }
        if ($this.Post -ne $null) {
            $attr += ('post="{0}"' -f $this.Post)
        }
        if ($this.TelNumber -ne $null) {
            $attr += ('telephone_number="{0}"' -f $this.TelNumber)
        }
        if ($this.Url -ne $null) {
            $attr += ('url="{0}"' -f $this.Url)
        }
        if ($this.Locale -ne $null) {
            $attr += ('locale="{0}"' -f $this.Locale)
        }
        if ($this.Base -ne $null) {
            $attr += ('base="{0}"' -f $this.Base)
        }
        if ($this.PrimaryGroup -ne $null) {
            $attr += ('primary_group="{0}"' -f $this.PrimaryGroup)
        }
        $elem = @()
        foreach ($org in $this.Organization) {
            if ($org -is [int]) {
                $elem += ('<organization>{0}</organization>' -f [string]$org)
            }
        }
        return [string]('<user_info {0}>{1}</user_info>' -f ($attr -join ' '), ($elem -join ''))
    }
}
Class GaroonAdmin : GaroonClass {
    [string] $ApiSuffix = "/sysapi/admin/api"
    GaroonAdmin() : base() {}
    GaroonAdmin([string]$URL) : base($URL) {}
    GaroonAdmin([string]$URL, [PSCredential] $Credential) : base($URL, $Credential) {}
    [int[]]GetUserIds([int]$Offset, [int]$Limit) {
        $Action = "AdminGetUserIds"
        $ParamBody = ('<parameters><offset xmlns="">{0}</offset><limit xmlns="">{1}</limit></parameters>' -f [string]$Offset, [string]$Limit)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.GetUserIdsResponse.returns.userId
    }
    [int[]]GetUserIds() {
        $Action = "AdminGetUserIds"
        $ParamBody = "<parameters></parameters>"
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.GetUserIdsResponse.returns.userId
    }
    [Object[]]GetUserDetailByIds([int[]]$UserId) {
        $Action = "AdminGetUserDetailByIds"
        [string[]]$body = $UserId | Foreach-Object {"<userId>{0}</userId>" -f $_}
        $ParamBody = ('<parameters xmlns="">{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.GetUserDetailByIdsResponse.returns.userDetail |
            ForEach-Object {[PSCustomObject]@{
                userId       = [int]$_.userId;
                login_name   = $_.login_name.Trim();
                display_name = $_.display_name.Trim()
            }
        }
    }
    [int]CountUsers() {
        $Action = "AdminCountUsers"
        $ParamBody = "<parameters></parameters>"
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.CountUsersResponse.returns.number_users
    }
    [int]CountUsersInOrg([int]$OrgId) {
        $Action = "AdminCountUsersInOrg"
        $ParamBody = ('<parameters><orgId>{0}</orgId></parameters>' -f [string]$OrgId)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.CountUsersInOrgResponse.returns.number_users
    }
    [int[]]GetUserIdsInOrg([int]$OrgId) {
        $Action = "AdminGetUserIdsInOrg"
        $ParamBody = ('<parameters><orgId>{0}</orgId></parameters>' -f [string]$OrgId)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.GetUserIdsInOrgResponse.returns.userId
    }
    [int]CountNoGroupUsers() {
        $Action = "AdminCountNoGroupUsers"
        $ParamBody = "<parameters></parameters>"
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.CountNoGroupUsersResponse.returns.number_users
    }
    [int[]]GetNoGroupUserIds() {
        $Action = "AdminGetNoGroupUserIds"
        $ParamBody = "<parameters></parameters>"
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.GetNoGroupUserIdsResponse.returns.userId
    }
    [int]CountOrgsOfUser([int]$UserId) {
        $Action = "AdminCountOrgsOfUser"
        $ParamBody = ('<parameters><userId>{0}</userId></parameters>' -f [string]$UserId)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.CountOrgsOfUserResponse.returns.number_orgs
    }
    [int[]]GetOrgIdsOfUser([int]$UserId) {
        $Action = "AdminGetOrgIdsOfUser"
        $ParamBody = ('<parameters><userId>{0}</userId></parameters>' -f [string]$UserId)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.GetOrgIdsOfUserResponse.returns.orgId
    }
    [int]GetUserIdByLoginName([string]$LoginName) {
        $Action = "AdminGetUserIdByLoginName"
        $ParamBody = ('<parameters><login_name>{0}</login_name></parameters>' -f $LoginName)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.GetUserIdByLoginNameResponse.returns.userId
    }
    [object[]]AddUserAccount([string]$LoginName, [string]$DisplayName, [string]$Password) {
        return $this.AddUserAccount($LoginName, $DisplayName, $Password, $null)
    }
    [object[]]AddUserAccount([string]$LoginName, [string]$DisplayName, [string]$Password, [Object]$UserInfo) {
        $Action = "AdminAddUserAccount"
        $_UserInfo = $null
        if ($UserInfo) {
            $_UserInfo = [UserInfo]::New($UserInfo).GetUserInfoString()
        }
        $body = @(
            ('<login_name xmlns="">{0}</login_name>' -f $LoginName),
            ('<display_name xmlns="">{0}</display_name>' -f $DisplayName),
            ('<password_raw xmlns="">{0}</password_raw>' -f $Password)
        )
        if ($_UserInfo) {
            $body += $_UserInfo
        }
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.AddUserAccountResponse.returns.userAccount |
            ForEach-Object {[PSCustomObject]@{
                userId       = [int]$_.userId;
                login_name   = $_.login_name.Trim();
                display_name = $_.display_name.Trim()
            }
        }
    }
    [int[]]RemoveUsersByIds([int[]]$UserId) {
        $Action = "AdminRemoveUsersByIds"
        [string[]]$body = $UserId | Foreach-Object {"<userId>{0}</userId>" -f $_}
        $ParamBody = ('<parameters xmlns="">{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.RemoveUsersByIdsResponse.returns.affected_userId
    }
    [object[]]ModifyUserAccount([int]$UserId, [string]$LoginName, [string]$DisplayName, [string]$Password, [Object]$UserInfo) {
        $_UserInfo = $null
        if ($UserInfo) {
            $_UserInfo = [UserInfo]::New($UserInfo).GetUserInfoString()
        }
        $Action = "AdminModifyUserAccount"
        $body = @(('<userId>{0}</userId>' -f [string]$UserId))
        if ($LoginName) { $body += ('<login_name xmlns="">{0}</login_name>' -f $LoginName) }
        if ($DisplayName) { $body += ('<display_name xmlns="">{0}</display_name>' -f $DisplayName) }
        if ($Password) { $body += ('<password_raw xmlns="">{0}</password_raw>' -f $Password) }
        if ($_UserInfo) { $body += $_UserInfo }
        if ($body.Length -eq 0) {
            return $null
        }
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.ModifyUserAccountResponse.returns.userAccount |
            ForEach-Object {[PSCustomObject]@{
                userId       = [int]$_.userId;
                login_name   = $_.login_name.Trim();
                display_name = $_.display_name.Trim()
            }
        }
    }
    [int[]]SetOrgsOfUser([int]$UserId, [int[]]$OrgId) {
        $Action = "AdminSetOrgsOfUser"
        [string[]]$body = $OrgId | Foreach-Object {"<orgId>{0}</orgId>" -f $_}
        $ParamBody = ('<parameters xmlns=""><userId>{0}</userId>{1}</parameters>' -f [string]$UserId, ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.SetOrgsOfUserResponse.returns.affected_orgId
    }
    [object[]]AddUsersToOrg([int]$OrgId, [int[]]$UserId) {
        $Action = "AdminAddUsersToOrg"
        $body = @(
            ('<orgId>{0}</orgId>' -f $OrgId)
        )
        $body += $UserId | Foreach-Object {"<userId>{0}</userId>" -f $_}
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.AddUsersToOrgResponse.returns |
            ForEach-Object {[PSCustomObject]@{
                number_relationships_affected = [int]$_.number_relationships_affected;
                affected_orgId                = [int]$_.affected_orgId;
                affected_userId               = [int[]]$_.affected_userId;
            }
        }
    }
    [object[]]AddOrg([string]$OrgCode, [string]$OrgName, [int]$ParentOrgId) {
        $Action = "AdminAddOrg"
        $body = @(
            ('<org_code>{0}</org_code>' -f $OrgCode),
            ('<org_name>{0}</org_name>' -f $OrgName)
        )
        if ($ParentOrgId -is [int]) {
            $body += ('<parent_orgId>{0}</parent_orgId>' -f [string]$ParentOrgId)
        }
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.AddOrgResponse.returns |
            ForEach-Object {[PSCustomObject]@{
                orgId        = [int]$_.org_info.orgId;
                org_code     = [string]$_.org_info.org_code.Trim();
                org_name     = [string]$_.org_info.org_name.Trim();
                parent_orgId = [int]$_.parent_orgId
            }
        }
    }
    [object[]]AddOrg([string]$OrgCode, [string]$OrgName) {
        return $this.AddOrg($OrgCode, $OrgName, $null)
    }
    [object[]]ModifyOrgInfo([int]$OrgId, [string]$OrgCode, [string]$OrgName) {
        $Action = "AdminModifyOrgInfo"
        $body = @(
            ('<orgId>{0}</orgId>' -f [string]$OrgId),
            ('<org_code>{0}</org_code>' -f $OrgCode),
            ('<org_name>{0}</org_name>' -f $OrgName)
        )
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.ModifyOrgInfoResponse.returns.org_info |
            ForEach-Object {[PSCustomObject]@{
                orgId    = [int]$_.orgId;
                org_code = [string]$_.org_code.Trim();
                org_name = [string]$_.org_name.Trim();
            }
        }
    }
    [int[]]RemoveOrgsByIds([int]$OrgId) {
        $Action = "AdminRemoveOrgsByIds"
        [string[]]$body = $OrgId | Foreach-Object {"<orgId>{0}</orgId>" -f $_}
        $ParamBody = ('<parameters xmlns="">{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.RemoveOrgsByIdsResponse.returns.affected_orgId
    }
    [object[]]RemoveUsersFromOrg([int]$OrgId, [int[]]$UserId) {
        $Action = "AdminRemoveUsersFromOrg"
        $body = @(
            ('<orgId>{0}</orgId>' -f $OrgId)
        )
        $body += $UserId | Foreach-Object {"<userId>{0}</userId>" -f $_}
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.RemoveUsersFromOrgResponse.returns |
            ForEach-Object {[PSCustomObject]@{
                number_relationships_affected = [int]$_.number_relationships_affected;
                affected_orgId                = [int]$_.affected_orgId;
                affected_userId               = [int[]]$_.affected_userId;
            }
        }
    }
    [object[]]AddChildrenOfOrg([int]$ParentOrgId, [int[]]$ChildOrgId) {
        $Action = "AdminAddChildrenOfOrg"
        $body = @(
            ('<parent_orgId>{0}</parent_orgId>' -f $ParentOrgId)
        )
        $body += $ChildOrgId | Foreach-Object {"<child_orgId>{0}</child_orgId>" -f $_}
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.AddChildrenOfOrgResponse.returns |
            ForEach-Object {[PSCustomObject]@{
                number_relationships_affected = [int]$_.number_relationships_affected;
                affected_parent_orgId         = [int]$_.affected_parent_orgId;
                affected_child_orgId          = [int[]]$_.affected_child_orgId;
            }
        }
    }
    [int]CountOrgs() {
        $Action = "AdminCountOrgs"
        $ParamBody = "<parameters></parameters>"
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.CountOrgsResponse.returns.number_orgs
    }
    [int[]]GetOrgIds() {
        $Action = "AdminGetOrgIds"
        $ParamBody = "<parameters></parameters>"
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.GetOrgIdsResponse.returns.orgId
    }
    [Object[]]GetOrgDetailByIds([int[]]$OrgId) {
        $Action = "AdminGetOrgDetailByIds"
        [string[]]$body = $OrgId | Foreach-Object {"<orgId>{0}</orgId>" -f $_}
        $ParamBody = ('<parameters xmlns="">{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.GetOrgDetailByIdsResponse.returns.orgDetail |
            ForEach-Object {[PSCustomObject]@{
                orgId    = [int]$_.orgId;
                org_code = $_.org_code.Trim();
                org_name = $_.org_name.Trim()
            }
        }
    }
    [int]CountChildOrgs([int]$ParentOrgId) {
        $Action = "AdminCountChildOrgs"
        $ParamBody = ('<parameters><parent_orgId>{0}</parent_orgId></parameters>' -f [string]$ParentOrgId)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.CountChildOrgsResponse.returns.number_child_orgs
    }
    [int[]]GetChildOrgs([int]$ParentOrgId) {
        $Action = "AdminGetChildOrgs"
        $ParamBody = ('<parameters><parent_orgId>{0}</parent_orgId></parameters>' -f [string]$ParentOrgId)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.GetChildOrgsResponse.returns.orgId
    }
    [int]GetParentOrgId([int]$ChildOrgId) {
        $Action = "AdminGetParentOrgId"
        $ParamBody = ('<parameters><child_orgId>{0}</child_orgId></parameters>' -f [string]$ChildOrgId)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.GetParentOrgIdResponse.returns.parent_orgId
    }
    [int]GetOrgIdByOrgCode([string]$OrgCode) {
        $Action = "AdminGetOrgIdByOrgCode"
        $ParamBody = ('<parameters><org_code>{0}</org_code></parameters>' -f $OrgCode)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.GetOrgIdByOrgCodeResponse.returns.orgId
    }
}
<#
.Synopsis
   ガルーンの組織を取得します
.DESCRIPTION
   組織名をキーにガルーンの組織を取得します。
   組織が見つからない場合、$nullを返します。
.PARAMETER OrganizationName
    取得したい組織の名前
    配列で複数渡すことができます
    パイプライン入力が可能です
    エイリアス：'Name', 'Organization'
.PARAMETER URL
    ガルーンのURL
    必ずトップページのURLを指定してください
    例: http://grnserver/cgi-bin/cgi/grn.cgi
.PARAMETER Credential
    ガルーンに接続するための資格情報
.PARAMETER SearchMode
    組織を検索する方法を指定します
    Equal:  完全一致検索（デフォルト）
    Like:   ワイルドカード検索
    RegExp: 正規表現検索
.PARAMETER NoDetail
    指定した場合、組織名、組織ID、組織コードのみを取得し、
    親子組織やメンバユーザの情報は取得しません。
    通信量を抑えて高速に結果を取得できます。
.EXAMPLE
    Get-GrnOrganization -OrganizationName '営業部' -URL 'http://grnserver/grn.cgi' -Credential (Get-Credential)
    Example 1: 組織名が"営業部"の組織の情報を取得する
    OrganizationName   : 営業部
    Code               : eigyo
    Id                 : 9
    ParentOrganization : 大洗株式会社
    ChildOrganization  : {第1営業グループ, 第2営業グループ}
    Members            : {nishizumi, takebe, reizen, akiyama...}
.EXAMPLE
    Get-GrnOrganization '*' -SearchMode Like -NoDetail -URL $GrnURL -Credential $PSCred
    Example 2: すべての組織を取得する（組織名と組織IDのみ）
    ワイルドカードを使用してガルーンに登録されたすべての組織を取得します。
    -NoDetail オプションを使用することで通信量を抑えて高速に取得できます。
.EXAMPLE
    Get-GrnOrganization "営業部","経理部" -URL $GrnURL -Credential $PSCred
    Example 3: 複数の組織を取得する
    組織名を配列で渡すことで複数組織の情報を取得できます。
    出力は入力した組織数と同数の要素を持つジャグ配列になります。
    ※注意： 組織名をパラメータではなくパイプラインで入力した場合は、ジャグ配列ではなくフラットな配列で出力されます
#>
function Get-GrnOrganization {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [Alias('Name', 'Organization')]
        [string[]]$OrganizationName,
        [Parameter(Mandatory = $true)]
        [string]$URL,
        [Parameter(Mandatory = $true)]
        [pscredential]$Credential,
        [ValidateSet("Equal", "Like", "RegExp")]
        [string]$SearchMode = "Equal",
        [switch]$NoDetail
    )
    Begin {
        $base = New-Object GaroonBase @($URL, $Credential) -ErrorAction Stop
        $admin = New-Object GaroonAdmin @($URL, $Credential) -ErrorAction Stop
        try {
            $orgids = $admin.GetOrgIds()
            if ($orgids.Count -le 0) {
                throw "組織情報が取得できませんでした"
            }
            else {
                $orgs = $base.GetOrganizationsById($orgids)
            }
        }
        catch [Exception] {
            Write-Error -Exception $_.Exception
            return
        }
        $private:ex = switch ($SearchMode) {
            'Equal' {'eq'}
            'Like' {'like'}
            'RegExp' {'match'}
        }
        Set-Variable -Name eval -Value ('$_.name -{0} $Org' -f $ex) -Option ReadOnly
    }
    Process {
        $Ret = @()
        foreach ($Org in $OrganizationName) {
            $private:s = $orgs.Where( {iex $eval})
            if ($s.Count -ge 1) {
                $Ret += , @($s | foreach {
                        if ($_.key) {
                            $OrgDetail = $admin.GetOrgDetailByIds($_.key)
                        }
                        if (-not $NoDetail) {
                            if ($_.organization) {
                                $ChildOrgs = $base.GetOrganizationsById($_.organization.key)
                            }
                            else {$ChildOrgs = $null}
                            if ($_.parent_organization) {
                                $ParentOrg = $base.GetOrganizationsById($_.parent_organization)
                            }
                            else {$ParentOrg = $null}
                            if ($_.members.user.id) {
                                $Members = $base.GetUsersById($_.members.user.id)
                            }
                            else {$Members = $null}
                        }
                        [PSCustomObject]@{
                            OrganizationName   = [string]$_.name    #組織名
                            Code               = [string]$OrgDetail.org_code    #組織コード
                            Id                 = [string]$_.key #組織ID
                            ParentOrganization = [string]$ParentOrg.name    #3親組織
                            ChildOrganization  = [string[]]$ChildOrgs.name  #子組織
                            Members            = [string[]]$Members.login_name  #メンバー
                        }
                    })
            }
            else {
                $Ret += , $null
            }
        }
        if (-not $Ret) {
            $null
        }
        elseif ($Ret.Count -eq 1) {
            $Ret[0]
        }
        else {
            $Ret
        }
        trap [Exception] {
            if ($_.Exception -is [System.Net.WebException]) {
                Write-Error -Exception $_.Exception
                return $null
            }
        }
    }
    End {
    }
}
function New-GrnOrganization {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [Alias('Name', 'Organization')]
        [string[]]$OrganizationName,
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Alias('Code')]
        [string]$OrganizationCode,
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Alias('Parent')]
        [string]$ParentOrganization,
        [Parameter(Mandatory = $true)]
        [string]$URL,
        [Parameter(Mandatory = $true)]
        [pscredential]$Credential,
        [switch]$PassThru
    )
    Begin {
        $admin = New-Object GaroonAdmin @($URL, $Credential) -ErrorAction Stop
    }
    Process {
        if (-not $OrganizationCode) {
            $OrganizationCode = -join ((1..9) | % {Get-Random -input ([char[]]((48..57) + (65..90) + (97..122)))})  #Gen random 9 chars passwd that only has 0-9A-Za-z
            Write-Warning "組織コードが指定されていません。自動生成された組織コード($OrganizationCode)を使用します"
        }
        if ($ParentOrganization) {
            $Parent = Get-GrnOrganization -OrganizationName $ParentOrganization -NoDetail -URL $URL -Credential $Credential -ErrorAction Stop
            if (-not $Parent) {
                Write-Error "親組織が見つかりません"
                return
            }
            elseif ($Parent.Count -ge 2) {
                Write-Error "指定の親組織名と同名の組織が複数見つかりました。処理を中止します"
                return
            }
        }
        try {
            [void] $admin.AddOrg($OrganizationCode, $OrganizationName, $Parent.Id)
        }
        catch {
            if ($_.Exception.Message -match 'GRN_CMMN_00103') {
                Write-Error '[ERROR][GRN_CMMN_00103] すでに存在する組織コードの組織を指定しています。'
            }
            else {
                Write-Error $_.Exception.Message
            }
        }
        if ($PassThru) {
            try {
                Get-GrnOrganization -OrganizationName $OrganizationName -URL $URL -Credential $Credential -ErrorAction Stop
            }
            catch {
                Write-Error $_.Exception.Message
            }
        }
    }
}
<#
.Synopsis
   ガルーンのユーザを取得します
.DESCRIPTION
   ログイン名をキーにガルーンのユーザを取得します。
   ユーザが見つからない場合、$nullを返します。
.EXAMPLE
    Get-GrnUser -LoginName 'sato' -URL 'http://grnserver/grn.cgi' -Credential (Get-Credential)
    LoginName           : sato
    Id                  : 6
    DisplayName         : 佐藤 昇
    Kana                : さとう のぼる
    Email               : sato@localhost
    Position            : 1
    PrimaryOrganization : 情報システム部
    Organization        : 情報システム部
    Invalid              : true
#>
function Get-GrnUser {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$LoginName,
        [Parameter(Mandatory = $true)]
        [string]$URL,
        [Parameter(Mandatory = $true)]
        [pscredential]$Credential
    )
    Begin {
        $base = New-Object GaroonBase @($URL, $Credential) -ErrorAction Stop
    }
    Process {
        try {
            $user = $base.GetUsersByLoginName($LoginName)
            if (-not $user.key) {
                return $null    #ユーザが見つからない
            }
            if ($user.organization.id) {
                $Orgs = $base.GetOrganizationsById($user.organization.id)
            }
            if ($user.primary_organization) {
                $PrimaryOrg = $base.GetOrganizationsById($user.primary_organization)
            }
            switch ($user.status) {
                0 { $Invalid = $false }   #Enable
                1 { $Invalid = $true}  #Disable
                2 { $Invalid = $true }  #Deleted
                Default {}
            }
            $Ret = [PSCustomObject]@{
                LoginName           = [string]$user.login_name
                Id                  = [string]$user.key
                DisplayName         = [string]$user.name
                Kana                = [string]$user.reading
                Email               = [string]$user.email
                Position            = [int]$user.order
                PrimaryOrganization = [string]$PrimaryOrg.name
                Organization        = [string[]]$Orgs.name
                Invalid             = [bool]$Invalid
                Description         = [string]$user.description
                Phone               = [string]$user.phone
                Post                = [string]$user.title
            }
            $Ret
        }
        catch [Exception] {
            Write-Error -Exception $_.Exception
            return $null
        }
    }
    End {}
}
<#
.Synopsis
   ガルーンにユーザを作成します
.DESCRIPTION
   ガルーンにユーザを作成します
.PARAMETER URL
    ガルーンのURL
    必ずトップページのURLを指定してください
    例: http://grnserver/cgi-bin/cgi/grn.cgi
.PARAMETER Credential
    ガルーンに接続するための資格情報
.PARAMETER PassThru
    デフォルトではこの関数は処理後に何も出力しません。
    -PassThruを使用すると作成したユーザ情報を出力します。
.PARAMETER LoginName
    ログイン名
.PARAMETER Password
    パスワード
.PARAMETER Position
    表示優先度
.PARAMETER Invalid
    使用状態(true:停止, false:使用)
.PARAMETER Kana
    よみがな
.PARAMETER Email
    メールアドレス
.PARAMETER Description
    メモ
.PARAMETER Post
    役職
.PARAMETER Phone
    電話番号
.PARAMETER PrimaryOrganization
    優先する組織の名称
.PARAMETER Organization
    所属する組織の名称（複数指定可）
    Organizationパラメータを指定するときは必ずPrimaryOrganizationパラメータも指定してください
    PrimaryOrganizationが指定されていない場合、Organizationパラメータの値は無視されます
.EXAMPLE
    New-GrnUser -LoginName 'yui' -DisplayName '平沢唯' -Password 'passw0rd' -URL $URL -Credential $cred
    Example 1: ログイン名が'yui'、表示名が'平沢唯'のユーザを作成します
    ログイン名、表示名、パスワードの3つが最低限必要なパラメータです
.EXAMPLE
    $UserInfo = @{
        LoginName = 'aoba'
        DisplayName = '涼風青葉'
        Password = 'Passw0rd'
        Kana = 'すずかぜあおば'
        Phone = '090-XXX-XXXX'
        PrimaryOrganization = 'イーグルジャンプ'
        Organization = ('イーグルジャンプ', '第一開発部', 'キャラ班')
    }
    PS C:\>New-GrnUser @UserInfo -URL $GrnURL -Credential $PSCred -PassThru
    Example 2: 詳細情報つきユーザ作成
    ユーザに設定する付加情報が多い場合、パラメータ設定に"スプラッティング"を使用すると便利です
#>
function New-GrnUser
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory)]
        [string]$URL,
        [Parameter(Mandatory)]
        [pscredential]$Credential,
        [switch]$PassThru,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)] [string]$LoginName, # ログイン名
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)] [string]$DisplayName, # 表示名
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)] [string]$Password, # パスワード
        [Parameter(ValueFromPipelineByPropertyName)] [int]$Position = -1, # 表示優先度
        [Parameter(ValueFromPipelineByPropertyName)] [bool]$Invalid = $false, # 使用状態(true:停止, false:使用)
        [Parameter(ValueFromPipelineByPropertyName)] [string]$Kana, # よみ
        [Parameter(ValueFromPipelineByPropertyName)] [string]$Email, # メールアドレス
        [Parameter(ValueFromPipelineByPropertyName)] [string]$Description, # メモ
        [Parameter(ValueFromPipelineByPropertyName)] [string]$Post, # 役職
        [Parameter(ValueFromPipelineByPropertyName)] [string]$Phone, # 電話番号
        [Parameter(ValueFromPipelineByPropertyName)] [string]$PrimaryOrganization, # 優先する組織
        [Parameter(ValueFromPipelineByPropertyName)] [string[]]$Organization, # 所属させる組織
        [parameter(DontShow, ValueFromPipelineByPropertyName)][string]$Id    # ユーザID（隠しパラメータ）
    )
    Begin {
        $base = New-Object GaroonBase @($URL, $Credential) -ErrorAction Stop
        $admin = New-Object GaroonAdmin @($URL, $Credential) -ErrorAction Stop
    }
    Process {
        if ((-not $PrimaryOrganization) -and ($Organization)) {
            Write-Warning ("優先する組織が指定されていません。このユーザはどの組織にも所属しません。")
            $Organization = $null
        }
        try {
            $user = $base.GetUsersByLoginName($LoginName)
        }
        catch [Exception] {
            if ($_.Exception -is [System.Net.WebException]) {
                Write-Error -Exception $_.Exception
                return
            }
        }
        if ($user.key) {
            $msg = ("ユーザ'{0}'は既に存在します" -f $LoginName)
            Write-Error $msg
            return
        }
        if (($PrimaryOrganization) -and ($Organization -notcontains $PrimaryOrganization)) {
            $Organization += $PrimaryOrganization
        }
        if ($Organization) {
            $Orgs = Get-GrnOrganization -OrganizationName $Organization -URL $URL -Credential $Credential -NoDetail
            if ($Orgs -contains $null) {
                $msg = ("所属組織一覧に存在しない組織名が含まれています")
                Write-Error $msg
                return
            }
            $local:flat = @($Orgs | foreach {$_} | sort Id -Unique)
            $local:P_OrgId = [int]$flat.Where( {$_.OrganizationName -eq $PrimaryOrganization}).Id
            $local:OrgIds = [int[]]$flat.Id
        }
        else {
            $local:P_OrgId = $null
            $local:OrgIds = $null
        }
        $UserInfo = @{
            PrimaryGroup = $P_OrgId
            Organization = $OrgIds
            Position     = $Position
            SortKey      = $Kana
            EmailAddress = $Email
            Description  = $Description
            Post         = $Post
            TelNumber    = $Phone
            Invalid      = $Invalid
        }
        try {
            [void] $admin.AddUserAccount($LoginName, $DisplayName, $Password, $UserInfo)
        }
        catch {
            Write-Error $_
        }
        if ($PassThru) {
            Get-GrnUser -LoginName $LoginName -URL $URL -Credential $Credential -ErrorAction SilentlyContinue
        }
    }
    End {}
}
<#
.Synopsis
   ガルーンのユーザ情報を変更します
.DESCRIPTION
   ガルーンのユーザ情報を変更します
   ログイン名(LoginName)がキーになります
.PARAMETER URL
    ガルーンのURL
    必ずトップページのURLを指定してください
    例: http://grnserver/cgi-bin/cgi/grn.cgi
.PARAMETER Credential
    ガルーンに接続するための資格情報
.PARAMETER PassThru
    デフォルトではこの関数は処理後に何も出力しません。
    -PassThruを使用すると作成したユーザ情報を出力します。
.PARAMETER LoginName
    変更するユーザのログイン名
    ログイン名をキーに変更ユーザを特定するため、ログイン名は変更できません
.PARAMETER Password
    パスワード
.PARAMETER Position
    表示優先度
.PARAMETER Invalid
    使用状態(true:停止, false:使用)
.PARAMETER Kana
    よみがな
.PARAMETER Email
    メールアドレス
.PARAMETER Description
    メモ
.PARAMETER Post
    役職
.PARAMETER Phone
    電話番号
.PARAMETER PrimaryOrganization
    優先する組織の名称
    Organizationを指定するときは必ずPrimaryOrganizationも指定してください
    また、PrimaryOrganizationを指定するときは必ずOrganizationも指定してください
.PARAMETER Organization
    所属する組織の名称（複数指定可）
    既存の組織に追加ではなく、上書きされることに注意してください
    組織を未所属にする場合、空文字列もしくは空配列を指定してください
.EXAMPLE
    Set-GrnUser -LoginName 'yui' -Password 'NewP@ssw0rd' -Email = 'new_yui@local' -URL $URL -Credential $cred
    Example 1: ログイン名が'yui'のユーザのパスワードとメールアドレスを変更します
.EXAMPLE
    Set-GrnUser -LoginName 'jun' -PrimaryOrganization = '軽音部' -Organization ('軽音部','ジャズ研') -URL $URL -Credential $cred
    Example 2: ログイン名が'jun'のユーザの所属組織を変更します
    組織を変更する場合、必ずPrimaryOrganizationパラメータを使用して優先する組織を指定します
.EXAMPLE
    Get-GrnUser -LoginName 'azusa' -U $u -C $c | Set-GrnUser -Organization @() -U $u -C $c -PassThru
    Example 3: パイプライン使用例
    パラメータをパイプラインで渡すことも可能です
    Organizationパラメータに空の配列を渡すことで、どの組織にも所属させないよう設定できます
#>
function Set-GrnUser
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory)]
        [string]$URL,
        [Parameter(Mandatory)]
        [pscredential]$Credential,
        [switch]$PassThru,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)] [string]$LoginName, # ログイン名
        [Parameter(ValueFromPipelineByPropertyName)] [string]$DisplayName, # 表示名
        [Parameter(ValueFromPipelineByPropertyName)] [string]$Password, # パスワード
        [Parameter(ValueFromPipelineByPropertyName)] [int]$Position, # 表示優先度
        [Parameter(ValueFromPipelineByPropertyName)] [bool]$Invalid, # 使用状態(true:停止, false:使用)
        [Parameter(ValueFromPipelineByPropertyName)] [string]$Kana, # よみ
        [Parameter(ValueFromPipelineByPropertyName)] [string]$Email, # メールアドレス
        [Parameter(ValueFromPipelineByPropertyName)] [string]$Description, # メモ
        [Parameter(ValueFromPipelineByPropertyName)] [string]$Post, # 役職
        [Parameter(ValueFromPipelineByPropertyName)] [string]$Phone, # 電話番号
        [Parameter(ValueFromPipelineByPropertyName)] [string]$PrimaryOrganization, # 優先する組織
        [Parameter(ValueFromPipelineByPropertyName)] [string[]]$Organization, # 所属させる組織
        [parameter(DontShow, ValueFromPipelineByPropertyName)][string]$Id    # ユーザID（隠しパラメータ）
    )
    Begin {
        $admin = New-Object GaroonAdmin @($URL, $Credential) -ErrorAction Stop
    }
    Process {
        $local:NoOrg = $false
        if (($PSBoundParameters.ContainsKey('Organization')) -and (-not $Organization)) {
            Write-Warning ("このユーザはどの組織にも所属しません。")
            $NoOrg = $true
        }
        if (-not $Id) {
            try {
                $current = Get-GrnUser -LoginName $LoginName -URL $URL -Credential $Credential -ErrorAction Stop
            }
            catch [Exception] {
                Write-Error -Exception $_.Exception
                return
            }
            ('DisplayName', 'Invalid', 'PrimaryOrganization', 'Organization', 'Position', 'Id').ForEach( {
                    if (-not $PSBoundParameters.ContainsKey($_)) {
                        Set-Variable $_ -Value $current.$_
                    }
                })
        }
        if ((-not $PrimaryOrganization) -and ($Organization)) {
            if (!$NoOrg) {Write-Warning ("優先する組織が指定されていません。このユーザはどの組織にも所属しません。")}
            $NoOrg = $true
        }
        if (($PrimaryOrganization) -and ($Organization -notcontains $PrimaryOrganization)) {
            $Organization += $PrimaryOrganization
        }
        if ($Organization -and !($NoOrg)) {
            $Orgs = Get-GrnOrganization -OrganizationName $Organization -URL $URL -Credential $Credential -NoDetail
            if ($Orgs -contains $null) {
                $msg = ("所属組織一覧に存在しない組織名が含まれています")
                Write-Error $msg
                return
            }
            $local:flat = @($Orgs | foreach {$_} | sort Id -Unique)
            $local:P_OrgId = [int]$flat.Where( {$_.OrganizationName -eq $PrimaryOrganization}).Id
            $local:OrgIds = [int[]]$flat.Id
        }
        else {
            $local:P_OrgId = $null
            $local:OrgIds = $null
        }
        $local:ParentParams = $PSBoundParameters
        $local:_tmp = {
            Param($_name)
            if ($ParentParams.ContainsKey($_name)) {
                Get-Variable $_name -ValueOnly
            }
            else {
                [NullString]::Value
            }
        }
        $UserInfo = @{
            PrimaryGroup = $P_OrgId
            Organization = $OrgIds
            Position     = $Position
            SortKey      = & $_tmp 'Kana'
            EmailAddress = & $_tmp 'Email'
            Description  = & $_tmp 'Description'
            Post         = & $_tmp 'Post'
            TelNumber    = & $_tmp 'Phone'
            Invalid      = $Invalid
        }
        try {
            [void] $admin.ModifyUserAccount($Id, $LoginName, $DisplayName, $Password, $UserInfo)
        }
        catch {
            Write-Error $_
        }
        if ($PassThru) {
            Get-GrnUser -LoginName $LoginName -URL $URL -Credential $Credential -ErrorAction SilentlyContinue
        }
    }
    End {}
}
<#
.Synopsis
   ガルーンのユーザを削除します
.DESCRIPTION
   指定したログイン名のガルーンユーザを削除します
.PARAMETER URL
    ガルーンのURL
    必ずトップページのURLを指定してください
    例: http://grnserver/cgi-bin/cgi/grn.cgi
.PARAMETER Credential
    ガルーンに接続するための資格情報
.PARAMETER LoginName
    削除するユーザのログイン名
.EXAMPLE
    Remove-GrnUser -LoginName 'miyamizu' -URL $URL -Credential $cred
    Example 1: ログイン名が'miyamizu'のユーザを削除します
#>
function Remove-GrnUser
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory)]
        [string]$URL,
        [Parameter(Mandatory)]
        [pscredential]$Credential,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)] [string]$LoginName, # ログイン名
        [parameter(DontShow, ValueFromPipelineByPropertyName)][int]$Id = -1    # ユーザID（隠しパラメータ）
    )
    Begin {
        $admin = New-Object GaroonAdmin @($URL, $Credential) -ErrorAction Stop
    }
    Process {
        try {
            $Id = $admin.GetUserIdByLoginName($LoginName)
            if ($Id -le 0) {
                throw ("ユーザ`"$LoginName`"が見つかりません")
            }
            else {
                Write-Verbose ("ユーザ`"$LoginName`"を削除します")
                $Removed = $admin.RemoveUsersByIds($Id)
                if ($Id -eq $Removed[0]) {
                    Write-Verbose "ユーザが削除されました"
                }
                else {
                    Write-Warning "Unknown Error"
                }
            }
        }
        catch [Exception] {
            Write-Error -Exception $_.Exception
            return
        }
    }
    End {}
}
function Invoke-SOAPRequest 
{
    [CmdletBinding()]
    [OutputType([Xml])]
    Param(
        [Xml]    $SOAPRequest, 
        [String] $URL 
    )
    Write-Verbose "Sending SOAP Request To Server: $URL" 
    $soapWebRequest = [System.Net.WebRequest]::Create($URL)
    $soapWebRequest.ContentType = 'text/xml;charset="utf-8"'
    $soapWebRequest.Accept      = "text/xml"
    $soapWebRequest.Method      = "POST"
    $requestStream = $soapWebRequest.GetRequestStream() 
    $SOAPRequest.Save($requestStream) 
    $requestStream.Close() 
    Write-Verbose "Send Complete, Waiting For Response." 
    $resp = $soapWebRequest.GetResponse() 
    $responseStream = $resp.GetResponseStream() 
    $soapReader = [System.IO.StreamReader]($responseStream) 
    $ReturnXml = [Xml] $soapReader.ReadToEnd() 
    $responseStream.Close() 
    Write-Verbose "Response Received."
    return $ReturnXml 
}
