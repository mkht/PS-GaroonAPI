using module ".\GaroonClass.ps1"

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

# メールAPI群実行用クラス
# https://cybozudev.zendesk.com/hc/ja/sections/200483090-%E3%83%A1%E3%83%BC%E3%83%AB
Class GaroonMail : GaroonClass {
    [string] $ApiSuffix = "/cbpapi/mail/api"
    GaroonMail() : base() {}
    GaroonMail([string]$URL) : base($URL) {}
    GaroonMail([string]$URL, [PSCredential] $Credential) : base($URL, $Credential) {}

    #メールアカウントを追加する
    [MailUserAccount[]]CreateUserAccount([Object[]]$MailUserAccount) {
        #入力を[MailUserAccount[]]にすると動かない…なぜ？
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

    #メールアカウントを更新する
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

    #メールアカウントを削除する
    [void]DeleteUserAccount([string]$AccountId, [bool]$DeleteAllEmail) {
        $Action = "MailDeleteUserAccount"
        $ParamBody = ('<parameters><delete_user_accounts xmlns="" account_id="{0}" delete_all_email="{1}"></delete_user_accounts></parameters>' -f $AccountId, $DeleteAllEmail.ToString().ToLower())
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return
    }

    #メールアカウントを取得する
    # API 実行ユーザー以外のユーザーのアカウントを取得することはできません。
    [Object[]]GetAccountsById([string[]]$AccountId) {
        $Action = "MailGetAccountsById"
        [string[]]$body = $AccountId | Foreach-Object {'<account_id xmlns="">{0}</account_id> ' -f $_}
        $ParamBody = ('<parameters>{0}</parameters>' -f ($body -join ''))
        $ReponseXml = $this.Request($this.CreateRequestXml($Action, $ParamBody, (Get-Date)))
        return $ReponseXml.Envelope.Body.MailGetAccountsByIdResponse.returns.account
    }
}