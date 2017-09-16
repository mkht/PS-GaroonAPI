using module PS-GaroonAPI

<#
.SYNOPSIS
    ガルーンの組織を作成します
.DESCRIPTION
    ガルーンの組織を作成します
.PARAMETER OrganizationName
    作成する組織の名前
    エイリアス: Name, Organization
.PARAMETER OrganizationCode
    組織コード
    他の組織の組織コードと重複しないコードを指定する必要があります。
    指定しない場合、ランダムな9文字の組織コードが自動で割り当てられます。
    エイリアス: Code
.PARAMETER ParentOrganization
    作成する組織の親組織の名前
    エイリアス: Parent
.PARAMETER URL
    ガルーンのURL
    必ずトップページのURLを指定してください
    例: http://grnserver/cgi-bin/cgi/grn.cgi
.PARAMETER Credential
    ガルーンに接続するための資格情報
.PARAMETER PassThru
    デフォルトではこの関数は処理後に何も出力しません。
    -PassThruを使用すると作成した組織情報を出力します。
.EXAMPLE
    New-GrnOrganization -OrganizationName '登山部' -URL $URL -Credential $cred
    Example 1: 組織名が'登山部'の組織を作成します
    親組織はなし、組織コードは自動割当
#>
function New-GrnOrganization {
    [CmdletBinding()]
    Param
    (
        # 追加する組織名
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [Alias('Name', 'Organization')]
        [string]$OrganizationName,

        #組織コード
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Alias('Code')]
        [string]$OrganizationCode,

        #親組織名
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Alias('Parent')]
        [string]$ParentOrganization,

        # ガルーンのURL
        [Parameter(Mandatory = $true)]
        [string]$URL,

        # ガルーン管理者の資格情報
        [Parameter(Mandatory = $true)]
        [pscredential]$Credential,

        [switch]$PassThru
    )

    Begin {
        $admin = New-Object GaroonAdmin @($URL, $Credential) -ErrorAction Stop
    }

    Process {
        #全組織情報取得
        # $AllOrg = Get-GrnOrganization -OrganizationName '*' -SearchMode Like -URL $URL -Credential $Credential -ErrorAction Stop
        # #組織コードチェック
        # if ($OrganizationCode -and ($OrganizationCode -cin $AllOrg.Code)) {
        #     Write-Error "指定された組織コードは既に使用されています"
        #     return
        # }
        # elseif (-not $OrganizationCode) {
        #     do {
        #         $OrganizationCode = [System.Web.Security.Membership]::GeneratePassword(8, 0)
        #         if ($OrganizationCode -cnotin $AllOrg.Code) {
        #             Write-Warning "組織コードが指定されていません。自動生成された組織コード($OrganizationCode)を使用します"
        #             break
        #         }
        #     } while ($true)
        # }

        # ↑↑↑↑ 組織コード重複チェックのために全組織取得をしていたが処理時間がかかりすぎるのでボツ
        # 組織コード重複の場合は実行時に[GRN_CMMN_00103]エラーが出るのでそれで。

        if (-not $OrganizationCode) {
            $OrganizationCode = -join ((1..9) | % {Get-Random -input ([char[]]((48..57) + (65..90) + (97..122)))})  #Gen random 9 chars passwd that only has 0-9A-Za-z
            Write-Warning "組織コードが指定されていません。自動生成された組織コード($OrganizationCode)を使用します"
        }

        #親組織チェック
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
            #組織作成
            [void] $admin.AddOrg($OrganizationCode, $OrganizationName, $Parent.Id)
        }
        catch {
            if ($_.Exception.Message -match 'GRN_CMMN_00103') {
                #標準のエラーメッセージだと分かりにくいのでメッセージをいじる
                #標準: [ERROR][GRN_CMMN_00103] 組織情報を設定できません。
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