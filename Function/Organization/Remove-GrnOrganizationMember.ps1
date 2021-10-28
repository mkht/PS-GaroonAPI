<#
.SYNOPSIS
    ガルーンの組織からユーザを削除します
.DESCRIPTION
    ガルーンの組織からユーザを削除します
.PARAMETER OrganizationName
    ユーザを削除する組織名
    エイリアス: Name, Organization
.PARAMETER Members
    削除するユーザのログインIDを配列で指定します
    パイプライン入力も可能です
.PARAMETER URL
    ガルーンのURL
    必ずトップページのURLを指定してください
    例: http://grnserver/cgi-bin/cgi/grn.cgi
.PARAMETER Credential
    ガルーンに接続するための資格情報
.PARAMETER PassThru
    デフォルトではこの関数は処理後に何も出力しません
    -PassThruを使用すると変更後の組織情報を出力します
.EXAMPLE
    Remove-GrnOrganizationMember -OrganizationName '星の浦女学院' -Members ('黒澤ダイヤ', '松浦果南') -URL $URL -Credential $cred
    Example 1: 組織「星の浦女学院」にユーザ「黒澤ダイヤ」と「松浦果南」を削除します
.NOTES
    削除するユーザは「ユーザ名」ではなく「ログインID」を指定することに注意してください
    削除するユーザが対象組織のメンバーではない場合でも、この関数は正常実行されエラーや警告は出力しません
#>
function Remove-GrnOrganizationMember {
    [CmdletBinding()]
    Param
    (
        # 対象組織名
        [Parameter(Mandatory = $true, ValueFromPipeline = $false, Position = 0)]
        [Alias('Name', 'Organization')]
        [string]$OrganizationName,

        # 削除するユーザ
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 1)]
        [string[]]$Members,

        # ガルーンのURL
        [Parameter(Mandatory = $true)]
        [string]$URL,

        # ガルーン管理者の資格情報
        [Parameter(Mandatory = $true)]
        [pscredential]$Credential,

        [switch]$PassThru
    )

    Begin {
        $base = New-Object GaroonBase @($URL, $Credential) -ErrorAction Stop
        $admin = New-Object GaroonAdmin @($URL, $Credential) -ErrorAction Stop

        $Org = Get-GrnOrganization $OrganizationName -NoDetail -URL $URL -Credential $Credential -wa SilentlyContinue
        if ((-not $Org) -or (@($Org).Length -ne 1)) {
            Write-Error ('組織 ({0}) が見つかりませんでした' -f $OrganizationName)
            return
        }
    }

    Process {
        #メンバー削除
        $RemoveUserIds = [int[]]($Members | ForEach-Object {
                $private:name = $_
                try { $private:user = $base.GetUsersByLoginName($name) }catch {}
                if ($user.key) {
                    $user.key
                }
                else {
                    Write-Warning ('指定されたログイン名のユーザ({0})が見つかりません' -f $name)
                }
            })
        if ($RemoveUserIds.Count -ge 1) {
            [void]$admin.RemoveUsersFromOrg($Org.Id, $RemoveUserIds)
        }
    }

    End {
        if ($PassThru) {
            Get-GrnOrganization $OrganizationName -URL $URL -Credential $Credential -ErrorAction Continue
        }
    }
}