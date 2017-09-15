using module PS-GaroonAPI

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
    出力は入力した組織数と同数の要素を持つ配列になります。
#>

function Get-GrnOrganization {
    [CmdletBinding()]
    Param
    (
        # 検索する組織名
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [Alias('Name', 'Organization')]
        [string[]]$OrganizationName,

        # ガルーンのURL
        [Parameter(Mandatory = $true)]
        [string]$URL,

        # ガルーン管理者の資格情報
        [Parameter(Mandatory = $true)]
        [pscredential]$Credential,

        #検索モード（完全一致、ワイルドカード、正規表現）
        [ValidateSet("Equal", "Like", "RegExp")]
        [string]$SearchMode = "Equal",

        #詳細情報を取得しない
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
        foreach ($Org in $OrganizationName) {
            $private:s = $orgs.Where( {iex $eval})
            if ($s.Count -ge 1) {
                $s | foreach {
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
                }
            }
            else {
                $null
            }
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