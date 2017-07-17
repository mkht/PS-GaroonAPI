using module PS-GaroonAPI

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
        # ガルーンのURL
        [Parameter(Mandatory)]
        [string]$URL,

        # ガルーン管理者の資格情報
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
        #[string]$Site,    # URL
        #[string]$Locale, # ロケールID
        #[string]$Base,   # 拠点ID
        [parameter(DontShow, ValueFromPipelineByPropertyName)][string]$Id    # ユーザID（隠しパラメータ）
    )

    Begin {
        $base = New-Object GaroonBase @($URL, $Credential) -ErrorAction Stop
        $admin = New-Object GaroonAdmin @($URL, $Credential) -ErrorAction Stop
    }
    Process {

        # 組織情報のチェック（「優先組織」が指定されていない場合、「所属組織」は無視する）
        if ((-not $PrimaryOrganization) -and ($Organization)) {
            Write-Warning ("優先する組織が指定されていません。このユーザはどの組織にも所属しません。")
            $Organization = $null
        }

        try {
            $user = $base.GetUsersByLoginName($LoginName)
        }
        catch [Exception] {
            # GRNERR以外のエラー(404など)の場合は中断
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