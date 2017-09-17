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

# システム管理API群実行用クラス
# https://cybozudev.zendesk.com/hc/ja/sections/200484594-%E3%82%B7%E3%82%B9%E3%83%86%E3%83%A0%E7%AE%A1%E7%90%86
Class GaroonAdmin : GaroonClass {
    [string] $ApiSuffix = "/sysapi/admin/api"
    GaroonAdmin() : base() {}
    GaroonAdmin([string]$URL) : base($URL) {}
    GaroonAdmin([string]$URL, [PSCredential] $Credential) : base($URL, $Credential) {}

    #ユーザーID一覧を取得する
    # 注意:"ユーザーID"はガルーン内のIDで、ログイン名とは異なる
    # Offset: 取得開始するユーザーのオフセット
    # Limit: 取得するユーザー数の上限
    [int[]]GetUserIds([int]$Offset, [int]$Limit) {
        $Action = "AdminGetUserIds"
        $ParamBody = ('<parameters><offset xmlns="">{0}</offset><limit xmlns="">{1}</limit></parameters>' -f [string]$Offset, [string]$Limit)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.GetUserIdsResponse.returns.userId
    }

    #ユーザーID一覧を取得する（全取得)
    # 注意:"ユーザーID"はガルーン内のIDで、ログイン名とは異なる
    [int[]]GetUserIds() {
        $Action = "AdminGetUserIds"
        $ParamBody = "<parameters></parameters>"
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.GetUserIdsResponse.returns.userId
    }

    #ユーザーIDからユーザー情報を取得する
    # 取得できるのはユーザID、ログイン名、表示名のみ。より詳細な情報が必要な場合はGaroonBase.GetUsersById()を使う
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

    #ユーザー数を取得する
    [int]CountUsers() {
        $Action = "AdminCountUsers"
        $ParamBody = "<parameters></parameters>"
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.CountUsersResponse.returns.number_users
    }

    #組織内のユーザー数を取得する
    [int]CountUsersInOrg([int]$OrgId) {
        $Action = "AdminCountUsersInOrg"
        $ParamBody = ('<parameters><orgId>{0}</orgId></parameters>' -f [string]$OrgId)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.CountUsersInOrgResponse.returns.number_users
    }

    #組織内のユーザーID一覧を取得する
    [int[]]GetUserIdsInOrg([int]$OrgId) {
        $Action = "AdminGetUserIdsInOrg"
        $ParamBody = ('<parameters><orgId>{0}</orgId></parameters>' -f [string]$OrgId)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.GetUserIdsInOrgResponse.returns.userId
    }

    #組織に未所属のユーザー数を取得する
    [int]CountNoGroupUsers() {
        $Action = "AdminCountNoGroupUsers"
        $ParamBody = "<parameters></parameters>"
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.CountNoGroupUsersResponse.returns.number_users
    }

    #組織に未所属のユーザーIDを取得する
    [int[]]GetNoGroupUserIds() {
        $Action = "AdminGetNoGroupUserIds"
        $ParamBody = "<parameters></parameters>"
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.GetNoGroupUserIdsResponse.returns.userId
    }

    #ユーザーが所属している組織数を取得する
    [int]CountOrgsOfUser([int]$UserId) {
        $Action = "AdminCountOrgsOfUser"
        $ParamBody = ('<parameters><userId>{0}</userId></parameters>' -f [string]$UserId)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.CountOrgsOfUserResponse.returns.number_orgs
    }

    #ユーザーが所属している組織のIDを取得する
    [int[]]GetOrgIdsOfUser([int]$UserId) {
        $Action = "AdminGetOrgIdsOfUser"
        $ParamBody = ('<parameters><userId>{0}</userId></parameters>' -f [string]$UserId)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.GetOrgIdsOfUserResponse.returns.orgId
    }

    #ログイン名からユーザーIDを取得する
    [int]GetUserIdByLoginName([string]$LoginName) {
        $Action = "AdminGetUserIdByLoginName"
        $ParamBody = ('<parameters><login_name>{0}</login_name></parameters>' -f $LoginName)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.GetUserIdByLoginNameResponse.returns.userId
    }

    #ユーザーを追加する
    [object[]]AddUserAccount([string]$LoginName, [string]$DisplayName, [string]$Password) {
        return $this.AddUserAccount($LoginName, $DisplayName, $Password, $null)
    }

    #ユーザーを追加する2
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

    #ユーザーを削除する
    [int[]]RemoveUsersByIds([int[]]$UserId) {
        $Action = "AdminRemoveUsersByIds"
        [string[]]$body = $UserId | Foreach-Object {"<userId>{0}</userId>" -f $_}
        $ParamBody = ('<parameters xmlns="">{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.RemoveUsersByIdsResponse.returns.affected_userId
    }

    #ユーザーを更新する
    # 注意1： https://support.cybozu.com/ja-jp/article/7948
    # 注意2： https://support.cybozu.com/ja-jp/article/7947
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

    #ユーザーを組織に所属させる
    [int[]]SetOrgsOfUser([int]$UserId, [int[]]$OrgId) {
        $Action = "AdminSetOrgsOfUser"
        [string[]]$body = $OrgId | Foreach-Object {"<orgId>{0}</orgId>" -f $_}
        $ParamBody = ('<parameters xmlns=""><userId>{0}</userId>{1}</parameters>' -f [string]$UserId, ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.SetOrgsOfUserResponse.returns.affected_orgId
    }

    #組織にユーザーを所属させる
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

    #組織を追加する
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

    #組織を追加する2
    [object[]]AddOrg([string]$OrgCode, [string]$OrgName) {
        return $this.AddOrg($OrgCode, $OrgName, $null)
    }

    #組織を更新する
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

    #組織を削除する
    [int[]]RemoveOrgsByIds([int]$OrgId) {
        $Action = "AdminRemoveOrgsByIds"
        [string[]]$body = $OrgId | Foreach-Object {"<orgId>{0}</orgId>" -f $_}
        $ParamBody = ('<parameters xmlns="">{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.RemoveOrgsByIdsResponse.returns.affected_orgId
    }

    #組織からユーザーを削除する
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

    #組織の親組織を変更する
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

    #組織数を取得する
    [int]CountOrgs() {
        $Action = "AdminCountOrgs"
        $ParamBody = "<parameters></parameters>"
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.CountOrgsResponse.returns.number_orgs
    }

    #組織IDを取得する（全取得)
    [int[]]GetOrgIds() {
        $Action = "AdminGetOrgIds"
        $ParamBody = "<parameters></parameters>"
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.GetOrgIdsResponse.returns.orgId
    }

    #組織IDから組織情報を取得する
    # より詳細な情報が必要な場合はGaroonBase.GetOrganizationsById()を使う
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

    #子組織の数を取得する
    [int]CountChildOrgs([int]$ParentOrgId) {
        $Action = "AdminCountChildOrgs"
        $ParamBody = ('<parameters><parent_orgId>{0}</parent_orgId></parameters>' -f [string]$ParentOrgId)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.CountChildOrgsResponse.returns.number_child_orgs
    }

    #子組織のIDを取得する
    [int[]]GetChildOrgs([int]$ParentOrgId) {
        $Action = "AdminGetChildOrgs"
        $ParamBody = ('<parameters><parent_orgId>{0}</parent_orgId></parameters>' -f [string]$ParentOrgId)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int[]]$ReponseXml.Envelope.Body.GetChildOrgsResponse.returns.orgId
    }

    #親組織のIDを取得する
    [int]GetParentOrgId([int]$ChildOrgId) {
        $Action = "AdminGetParentOrgId"
        $ParamBody = ('<parameters><child_orgId>{0}</child_orgId></parameters>' -f [string]$ChildOrgId)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.GetParentOrgIdResponse.returns.parent_orgId
    }

    #組織コードから組織IDを取得する
    [int]GetOrgIdByOrgCode([string]$OrgCode) {
        $Action = "AdminGetOrgIdByOrgCode"
        $ParamBody = ('<parameters><org_code>{0}</org_code></parameters>' -f $OrgCode)
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return [int]$ReponseXml.Envelope.Body.GetOrgIdByOrgCodeResponse.returns.orgId
    }


}
