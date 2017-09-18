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
        # ガルーンのURL
        [Parameter(Mandatory)]
        [string]$URL,

        # ガルーン管理者の資格情報
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
        #[string]$Site,    # URL
        #[string]$Locale, # ロケールID
        #[string]$Base,   # 拠点ID
        [parameter(DontShow, ValueFromPipelineByPropertyName)][string]$Id    # ユーザID（隠しパラメータ）
    )

    Begin {
        $admin = New-Object GaroonAdmin @($URL, $Credential) -ErrorAction Stop

    }
    Process {
        $local:NoOrg = $false

        # Organizationに空配列が渡された場合は無所属ユーザにする
        if (($PSBoundParameters.ContainsKey('Organization')) -and (-not $Organization)) {
            Write-Warning ("このユーザはどの組織にも所属しません。")
            $NoOrg = $true
        }

        # 既存のユーザ情報を取得(パイプラインでパラメータを渡された場合はスキップ)
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

        # Organizationが指定されているがPrimaryOrganizationが未指定の場合は未所属とする
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
