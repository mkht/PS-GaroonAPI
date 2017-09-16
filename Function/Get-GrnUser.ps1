using module PS-GaroonAPI

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
        # 検索するログイン名
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [string]$LoginName,

        # ガルーンのURL
        [Parameter(Mandatory = $true)]
        [string]$URL,

        # ガルーン管理者の資格情報
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
                Write-Warning ('指定されたログイン名のユーザが見つかりません')
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