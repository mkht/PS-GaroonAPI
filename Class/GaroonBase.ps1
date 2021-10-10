# ベースAPI群実行用クラス
# https://cybozudev.zendesk.com/hc/ja/sections/200483120-%E3%83%99%E3%83%BC%E3%82%B9
Class GaroonBase : GaroonClass {
    [string] $ApiSuffix = "/cbpapi/base/api"
    GaroonBase() : base() {}
    GaroonBase([string]$URL) : base($URL) {}
    GaroonBase([string]$URL, [PSCredential] $Credential) : base($URL, $Credential) {}

    #アプリケーションの使用状況を取得する
    [Object[]]GetApplicationStatus() {
        $Action = "BaseGetApplicationStatus"
        $ParamBody = '<parameters></parameters>'
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.BaseGetApplicationStatusResponse.returns.application
    }

    #ユーザーIDからユーザーを取得する
    # 注意:"ユーザーID"はガルーン内のIDで、ログイン名とは異なる
    [Object[]]GetUsersById([int[]]$UserId) {
        $Action = "BaseGetUsersById"
        [string[]]$body = $UserId | Foreach-Object { "<user_id>{0}</user_id>" -f $_ }
        $ParamBody = ('<parameters xmlns="">{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.BaseGetUsersByIdResponse.returns.user
    }

    #ログイン名からユーザーを取得する
    [Object[]]GetUsersByLoginName([string[]]$LoginName) {
        $Action = "BaseGetUsersByLoginName"
        [string[]]$body = $LoginName | Foreach-Object { "<login_name>{0}</login_name>" -f $_ }
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.BaseGetUsersByLoginNameResponse.returns.user
    }

    #組織IDから組織情報を取得する
    [Object[]]GetOrganizationsById([int[]]$OrganizationId) {
        $Action = "BaseGetOrganizationsById"
        [string[]]$body = $OrganizationId | Foreach-Object { "<organization_id>{0}</organization_id>" -f $_ }
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ResponseXml.Envelope.Body.BaseGetOrganizationsByIdResponse.returns.organization
    }
}