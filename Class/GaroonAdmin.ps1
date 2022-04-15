using namespace System.Xml
using namespace System.Security

# ユーザー情報の詳細を表す型
# https://cybozudev.zendesk.com/hc/ja/articles/202496120#step3
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
    Hidden [XmlDocument]$XmlDoc

    <# ---- コンストラクタ ---- #>
    UserInfo() {
        $this.XmlDoc = New-Object XmlDocument
        $this.XmlDoc.PreserveWhitespace = $true
    }

    UserInfo([Object]$Info) {
        $this.XmlDoc = New-Object XmlDocument
        $this.XmlDoc.PreserveWhitespace = $true

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

    [XmlElement]GetUserInfoXmlElement() {
        $userInfo = $this.XmlDoc.CreateElement('user_info')
        if ($this.Position -as [System.UInt32]) {
            $userInfo.SetAttribute('position', [string]$this.Position)
        }
        if ($this.Invalid -is [bool]) {
            $userInfo.SetAttribute('invalid', ([string]$this.Invalid).ToLower())
        }
        if ($null -ne $this.SortKey) {
            $userInfo.SetAttribute('sort_key', $this.SortKey)
        }
        if ($null -ne $this.EmailAddress) {
            $userInfo.SetAttribute('email_address', $this.EmailAddress)
        }
        if ($null -ne $this.Description) {
            $userInfo.SetAttribute('description', $this.Description)
        }
        if ($null -ne $this.Post) {
            $userInfo.SetAttribute('post', $this.Post)
        }
        if ($null -ne $this.TelNumber) {
            $userInfo.SetAttribute('telephone_number', $this.TelNumber)
        }
        if ($null -ne $this.Url) {
            $userInfo.SetAttribute('url', $this.Url)
        }
        if ($null -ne $this.Locale) {
            $userInfo.SetAttribute('locale', $this.Locale)
        }
        if ($null -ne $this.Base) {
            $userInfo.SetAttribute('base', $this.Base)
        }
        if ($null -ne $this.PrimaryGroup) {
            $userInfo.SetAttribute('primary_group', $this.PrimaryGroup)
        }

        $elem = @()
        foreach ($org in $this.Organization) {
            if ($org -is [int]) {
                $elem = $this.XmlDoc.CreateElement('organization')
                $elem.InnerText = [string]$org
                $userInfo.AppendChild($elem)
            }
        }

        return $userInfo
    }

    [string]GetUserInfoString() {
        return $this.GetUserInfoXmlElement().OuterXml
    }
}

# システム管理API群実行用クラス
# https://cybozudev.zendesk.com/hc/ja/sections/200484594-%E3%82%B7%E3%82%B9%E3%83%86%E3%83%A0%E7%AE%A1%E7%90%86
Class GaroonAdmin : GaroonClass {
    [string] $ApiSuffix = '/sysapi/admin/api'
    GaroonAdmin() : base() {}
    GaroonAdmin([string]$URL) : base($URL) {}
    GaroonAdmin([string]$URL, [PSCredential] $Credential) : base($URL, $Credential) {}

    #ユーザーID一覧を取得する
    # 注意:"ユーザーID"はガルーン内のIDで、ログイン名とは異なる
    # Offset: 取得開始するユーザーのオフセット
    # Limit: 取得するユーザー数の上限
    [int[]]GetUserIds([int]$Offset, [int]$Limit) {
        $Action = 'AdminGetUserIds'
        $ParamBody = ('<parameters><offset xmlns="">{0}</offset><limit xmlns="">{1}</limit></parameters>' -f [string]$Offset, [string]$Limit)
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ResponseXml.Envelope.Body.GetUserIdsResponse.returns.userId
    }

    #ユーザーID一覧を取得する（全取得)
    # 注意:"ユーザーID"はガルーン内のIDで、ログイン名とは異なる
    [int[]]GetUserIds() {
        $Action = 'AdminGetUserIds'
        $ParamBody = '<parameters></parameters>'
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ResponseXml.Envelope.Body.GetUserIdsResponse.returns.userId
    }

    #ユーザーIDからユーザー情報を取得する
    # 取得できるのはユーザID、ログイン名、表示名のみ。より詳細な情報が必要な場合はGaroonBase.GetUsersById()を使う
    [Object[]]GetUserDetailByIds([int[]]$UserId) {
        $Action = 'AdminGetUserDetailByIds'
        [string[]]$body = $UserId | ForEach-Object { '<userId>{0}</userId>' -f $_ }
        $ParamBody = ('<parameters xmlns="">{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.GetUserDetailByIdsResponse.returns.userDetail |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [PSCustomObject]@{
                userId       = [int]$_.userId;
                login_name   = $_.login_name.Trim();
                display_name = $_.display_name.Trim()
            }
        }
    }

    #ユーザー数を取得する
    [int]CountUsers() {
        $Action = 'AdminCountUsers'
        $ParamBody = '<parameters></parameters>'
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ResponseXml.Envelope.Body.CountUsersResponse.returns.number_users
    }

    #組織内のユーザー数を取得する
    [int]CountUsersInOrg([int]$OrgId) {
        $Action = 'AdminCountUsersInOrg'
        $ParamBody = ('<parameters><orgId>{0}</orgId></parameters>' -f [string]$OrgId)
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ResponseXml.Envelope.Body.CountUsersInOrgResponse.returns.number_users
    }

    #組織内のユーザーID一覧を取得する
    [int[]]GetUserIdsInOrg([int]$OrgId) {
        $Action = 'AdminGetUserIdsInOrg'
        $ParamBody = ('<parameters><orgId>{0}</orgId></parameters>' -f [string]$OrgId)
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ResponseXml.Envelope.Body.GetUserIdsInOrgResponse.returns.userId
    }

    #組織に未所属のユーザー数を取得する
    [int]CountNoGroupUsers() {
        $Action = 'AdminCountNoGroupUsers'
        $ParamBody = '<parameters></parameters>'
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ResponseXml.Envelope.Body.CountNoGroupUsersResponse.returns.number_users
    }

    #組織に未所属のユーザーIDを取得する
    [int[]]GetNoGroupUserIds() {
        $Action = 'AdminGetNoGroupUserIds'
        $ParamBody = '<parameters></parameters>'
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ResponseXml.Envelope.Body.GetNoGroupUserIdsResponse.returns.userId
    }

    #ユーザーが所属している組織数を取得する
    [int]CountOrgsOfUser([int]$UserId) {
        $Action = 'AdminCountOrgsOfUser'
        $ParamBody = ('<parameters><userId>{0}</userId></parameters>' -f [string]$UserId)
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ResponseXml.Envelope.Body.CountOrgsOfUserResponse.returns.number_orgs
    }

    #ユーザーが所属している組織のIDを取得する
    [int[]]GetOrgIdsOfUser([int]$UserId) {
        $Action = 'AdminGetOrgIdsOfUser'
        $ParamBody = ('<parameters><userId>{0}</userId></parameters>' -f [string]$UserId)
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ResponseXml.Envelope.Body.GetOrgIdsOfUserResponse.returns.orgId
    }

    #ログイン名からユーザーIDを取得する
    [int]GetUserIdByLoginName([string]$LoginName) {
        $Action = 'AdminGetUserIdByLoginName'
        $ParamBody = ('<parameters><login_name>{0}</login_name></parameters>' -f [SecurityElement]::Escape($LoginName))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ResponseXml.Envelope.Body.GetUserIdByLoginNameResponse.returns.userId
    }

    #ユーザーを追加する
    [object[]]AddUserAccount([string]$LoginName, [string]$DisplayName, [string]$Password) {
        return $this.AddUserAccount($LoginName, $DisplayName, $Password, $null)
    }

    #ユーザーを追加する2
    [object[]]AddUserAccount([string]$LoginName, [string]$DisplayName, [string]$Password, [Object]$UserInfo) {
        $Action = 'AdminAddUserAccount'

        $_UserInfo = $null
        if ($UserInfo) {
            $_UserInfo = [UserInfo]::New($UserInfo).GetUserInfoString()
        }

        $body = @(
            ('<login_name xmlns="">{0}</login_name>' -f [SecurityElement]::Escape($LoginName)),
            ('<display_name xmlns="">{0}</display_name>' -f [SecurityElement]::Escape($LoginName)),
            ('<password_raw xmlns="">{0}</password_raw>' -f [SecurityElement]::Escape($LoginName))
        )
        if ($_UserInfo) {
            $body += $_UserInfo
        }

        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.AddUserAccountResponse.returns.userAccount |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [PSCustomObject]@{
                userId       = [int]$_.userId;
                login_name   = $_.login_name.Trim();
                display_name = $_.display_name.Trim()
            }
        }
    }

    #ユーザーを削除する
    [int[]]RemoveUsersByIds([int[]]$UserId) {
        $Action = 'AdminRemoveUsersByIds'
        [string[]]$body = $UserId | ForEach-Object { '<userId>{0}</userId>' -f $_ }
        $ParamBody = ('<parameters xmlns="">{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ResponseXml.Envelope.Body.RemoveUsersByIdsResponse.returns.affected_userId
    }

    #ユーザーを更新する
    # 注意1： https://support.cybozu.com/ja-jp/article/7948
    # 注意2： https://support.cybozu.com/ja-jp/article/7947
    [object[]]ModifyUserAccount([int]$UserId, [string]$LoginName, [string]$DisplayName, [string]$Password, [Object]$UserInfo) {
        $_UserInfo = $null
        if ($UserInfo) {
            $_UserInfo = [UserInfo]::New($UserInfo).GetUserInfoString()
        }
        $Action = 'AdminModifyUserAccount'

        $body = @(('<userId>{0}</userId>' -f [string]$UserId))
        if ($LoginName) { $body += ('<login_name xmlns="">{0}</login_name>' -f [SecurityElement]::Escape($LoginName)) }
        if ($DisplayName) { $body += ('<display_name xmlns="">{0}</display_name>' -f [SecurityElement]::Escape($DisplayName)) }
        if ($Password) { $body += ('<password_raw xmlns="">{0}</password_raw>' -f [SecurityElement]::Escape($Password)) }
        if ($_UserInfo) { $body += $_UserInfo }
        if ($body.Length -eq 0) {
            return $null
        }

        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.ModifyUserAccountResponse.returns.userAccount |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [PSCustomObject]@{
                userId       = [int]$_.userId;
                login_name   = $_.login_name.Trim();
                display_name = $_.display_name.Trim()
            }
        }
    }

    #ユーザーを組織に所属させる
    [int[]]SetOrgsOfUser([int]$UserId, [int[]]$OrgId) {
        $Action = 'AdminSetOrgsOfUser'
        [string[]]$body = $OrgId | ForEach-Object { '<orgId>{0}</orgId>' -f $_ }
        $ParamBody = ('<parameters xmlns=""><userId>{0}</userId>{1}</parameters>' -f [string]$UserId, ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ResponseXml.Envelope.Body.SetOrgsOfUserResponse.returns.affected_orgId
    }

    #組織にユーザーを所属させる
    [object[]]AddUsersToOrg([int]$OrgId, [int[]]$UserId) {
        $Action = 'AdminAddUsersToOrg'
        $body = @(
            ('<orgId>{0}</orgId>' -f $OrgId)
        )
        $body += $UserId | ForEach-Object { '<userId>{0}</userId>' -f $_ }
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.AddUsersToOrgResponse.returns |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [PSCustomObject]@{
                number_relationships_affected = [int]$_.number_relationships_affected;
                affected_orgId                = [int]$_.affected_orgId;
                affected_userId               = [int[]]$_.affected_userId;
            }
        }
    }

    #組織を追加する
    [object[]]AddOrg([string]$OrgCode, [string]$OrgName, [int]$ParentOrgId) {
        $Action = 'AdminAddOrg'
        $body = @(
            ('<org_code>{0}</org_code>' -f [SecurityElement]::Escape($OrgCode)),
            ('<org_name>{0}</org_name>' -f [SecurityElement]::Escape($OrgName))
        )
        if ($ParentOrgId -is [int]) {
            $body += ('<parent_orgId>{0}</parent_orgId>' -f [string]$ParentOrgId)
        }

        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.AddOrgResponse.returns |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [PSCustomObject]@{
                orgId        = [int]$_.org_info.orgId;
                org_code     = [string]$_.org_info.org_code.Trim();
                org_name     = [string]$_.org_info.org_name.Trim();
                parent_orgId = [int]$_.parent_orgId
            }
        }
    }

    #組織を追加する2
    [object[]]AddOrg([string]$OrgCode, [string]$OrgName) {
        return $this.AddOrg($OrgCode, $OrgName, $null)
    }

    #組織を更新する
    [object[]]ModifyOrgInfo([int]$OrgId, [string]$OrgCode, [string]$OrgName) {
        $Action = 'AdminModifyOrgInfo'
        $body = @(
            ('<orgId>{0}</orgId>' -f [string]$OrgId),
            ('<org_code>{0}</org_code>' -f [SecurityElement]::Escape($OrgCode)),
            ('<org_name>{0}</org_name>' -f [SecurityElement]::Escape($OrgName))
        )
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.ModifyOrgInfoResponse.returns.org_info |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [PSCustomObject]@{
                orgId    = [int]$_.orgId;
                org_code = [string]$_.org_code.Trim();
                org_name = [string]$_.org_name.Trim();
            }
        }
    }

    #組織を削除する
    [int[]]RemoveOrgsByIds([int]$OrgId) {
        $Action = 'AdminRemoveOrgsByIds'
        [string[]]$body = $OrgId | ForEach-Object { '<orgId>{0}</orgId>' -f $_ }
        $ParamBody = ('<parameters xmlns="">{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ResponseXml.Envelope.Body.RemoveOrgsByIdsResponse.returns.affected_orgId
    }

    #組織からユーザーを削除する
    [object[]]RemoveUsersFromOrg([int]$OrgId, [int[]]$UserId) {
        $Action = 'AdminRemoveUsersFromOrg'
        $body = @(
            ('<orgId>{0}</orgId>' -f $OrgId)
        )
        $body += $UserId | ForEach-Object { '<userId>{0}</userId>' -f $_ }
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.RemoveUsersFromOrgResponse.returns |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [PSCustomObject]@{
                number_relationships_affected = [int]$_.number_relationships_affected;
                affected_orgId                = [int]$_.affected_orgId;
                affected_userId               = [int[]]$_.affected_userId;
            }
        }
    }

    #組織の親組織を変更する
    [object[]]AddChildrenOfOrg([int]$ParentOrgId, [int[]]$ChildOrgId) {
        $Action = 'AdminAddChildrenOfOrg'
        $body = @(
            ('<parent_orgId>{0}</parent_orgId>' -f $ParentOrgId)
        )
        $body += $ChildOrgId | ForEach-Object { '<child_orgId>{0}</child_orgId>' -f $_ }
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.AddChildrenOfOrgResponse.returns |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [PSCustomObject]@{
                number_relationships_affected = [int]$_.number_relationships_affected;
                affected_parent_orgId         = [int]$_.affected_parent_orgId;
                affected_child_orgId          = [int[]]$_.affected_child_orgId;
            }
        }
    }

    #組織数を取得する
    [int]CountOrgs() {
        $Action = 'AdminCountOrgs'
        $ParamBody = '<parameters></parameters>'
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ResponseXml.Envelope.Body.CountOrgsResponse.returns.number_orgs
    }

    #組織IDを取得する（全取得)
    [int[]]GetOrgIds() {
        $Action = 'AdminGetOrgIds'
        $ParamBody = '<parameters></parameters>'
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ResponseXml.Envelope.Body.GetOrgIdsResponse.returns.orgId
    }

    #組織IDから組織情報を取得する
    # より詳細な情報が必要な場合はGaroonBase.GetOrganizationsById()を使う
    [Object[]]GetOrgDetailByIds([int[]]$OrgId) {
        $Action = 'AdminGetOrgDetailByIds'
        [string[]]$body = $OrgId | ForEach-Object { '<orgId>{0}</orgId>' -f $_ }
        $ParamBody = ('<parameters xmlns="">{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.GetOrgDetailByIdsResponse.returns.orgDetail |`
            Where-Object { $null -ne $_ } |`
            ForEach-Object { [PSCustomObject]@{
                orgId    = [int]$_.orgId;
                org_code = $_.org_code.Trim();
                org_name = $_.org_name.Trim()
            }
        }
    }

    #子組織の数を取得する
    [int]CountChildOrgs([int]$ParentOrgId) {
        $Action = 'AdminCountChildOrgs'
        $ParamBody = ('<parameters><parent_orgId>{0}</parent_orgId></parameters>' -f [string]$ParentOrgId)
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ResponseXml.Envelope.Body.CountChildOrgsResponse.returns.number_child_orgs
    }

    #子組織のIDを取得する
    [int[]]GetChildOrgs([int]$ParentOrgId) {
        $Action = 'AdminGetChildOrgs'
        $ParamBody = ('<parameters><parent_orgId>{0}</parent_orgId></parameters>' -f [string]$ParentOrgId)
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ResponseXml.Envelope.Body.GetChildOrgsResponse.returns.orgId
    }

    #親組織のIDを取得する
    [int]GetParentOrgId([int]$ChildOrgId) {
        $Action = 'AdminGetParentOrgId'
        $ParamBody = ('<parameters><child_orgId>{0}</child_orgId></parameters>' -f [string]$ChildOrgId)
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ResponseXml.Envelope.Body.GetParentOrgIdResponse.returns.parent_orgId
    }

    #組織コードから組織IDを取得する
    [int]GetOrgIdByOrgCode([string]$OrgCode) {
        $Action = 'AdminGetOrgIdByOrgCode'
        $ParamBody = ('<parameters><org_code>{0}</org_code></parameters>' -f [SecurityElement]::Escape($OrgCode))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ResponseXml.Envelope.Body.GetOrgIdByOrgCodeResponse.returns.orgId
    }
}
