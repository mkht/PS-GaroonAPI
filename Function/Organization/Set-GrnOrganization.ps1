﻿<#
.SYNOPSIS
    ガルーンの組織情報を変更します
.DESCRIPTION
    ガルーンの組織情報を変更します
.PARAMETER OrganizationName
    変更する組織の名前
    エイリアス: Name, Organization
.PARAMETER NewOrganizationName
    組織名を変更する場合に、変更後の組織名を指定します
    エイリアス: NewName
.PARAMETER ParentOrganization
    新しい親組織の名前
    GaroonAPIの制限により親組織を削除することはできません
    エイリアス: Parent
.PARAMETER Members
    組織に所属するユーザのログインIDを配列で指定します
    既存の所属ユーザは削除されます
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
    Set-GrnOrganization -OrganizationName 'ゴルフ同好会' -NewOrganizationName 'ゴルフ部' -URL $URL -Credential $cred
    Example 1: 'ゴルフ同好会'の組織名を'ゴルフ部'に変更します
.EXAMPLE
    Set-GrnOrganization -OrganizationName 'テニス部' -ParentOrganization '運動部' -URL $URL -Credential $cred
    Example 2: 'テニス部'の親組織を'運動部'に変更します
.EXAMPLE
    Set-GrnOrganization -OrganizationName '水泳部' -Members ('nanase', 'tachibana', 'matsuoka') -URL $URL -Credential $cred
    Example 3: '水泳部'の所属メンバーを変更します
#>
function Set-GrnOrganization {
    [CmdletBinding()]
    Param
    (
        # 変更する組織名
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [Alias('Name', 'Organization')]
        [string]$OrganizationName,

        # 新しい組織名
        [Parameter()]
        [Alias('NewName')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( { $_ -ne $OrganizationName })]
        [string]$NewOrganizationName,

        # 新しい親組織名
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( { $_ -ne $OrganizationName })]
        [Alias('Parent')]
        [string]$ParentOrganization,

        # メンバー
        [Parameter()]
        [AllowEmptyCollection()]
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
    }

    Process {
        if ($PSBoundParameters.Members -is [Array]) {
            $Org = Get-GrnOrganization $OrganizationName -URL $URL -Credential $Credential -wa SilentlyContinue
        }
        else {
            $Org = Get-GrnOrganization $OrganizationName -NoDetail -URL $URL -Credential $Credential -wa SilentlyContinue
        }
        if ((-not $Org) -or (@($Org).Length -ne 1)) {
            Write-Error ('組織 ({0}) が見つかりませんでした' -f $OrganizationName)
            return
        }

        # 親組織の変更
        if ($PSBoundParameters.ParentOrganization) {
            if ($Org.ParentOrganization -eq $ParentOrganization) {
                Write-Host ('{0} の親組織は既に {1} になっています' -f $OrganizationName, $ParentOrganization)
            }
            else {
                $ParentOrg = Get-GrnOrganization $ParentOrganization -NoDetail -URL $URL -Credential $Credential -ea SilentlyContinue -wa SilentlyContinue
                if ((-not $ParentOrg) -or (@($ParentOrg).Length -ne 1)) {
                    Write-Error ('新しい親組織 ({0}) が見つかりませんでした' -f $ParentOrganization)
                }
                else {
                    [void]$admin.AddChildrenOfOrg($ParentOrg.Id, $Org.Id)
                }
            }
        }

        if ($PSBoundParameters.Members -is [Array]) {
            # 全メンバーを削除
            $RemoveUserIds = [int[]]($Org.Members | ForEach-Object {
                    try { $private:user = $base.GetUsersByLoginName($_) }catch {}
                    if ($user.key) {
                        $user.key
                    }
                })
            if ($RemoveUserIds.Count -ge 1) {
                [void]$admin.RemoveUsersFromOrg($Org.Id, $RemoveUserIds)
            }

            #メンバー追加
            $AddUserIds = [int[]]($Members | ForEach-Object {
                    $private:name = $_
                    try { $private:user = $base.GetUsersByLoginName($name) }catch {}
                    if ($user.key) {
                        $user.key
                    }
                    else {
                        Write-Warning ('指定されたログイン名のユーザ({0})が見つかりません' -f $name)
                    }
                })
            if ($AddUserIds.Count -ge 1) {
                [void]$admin.AddUsersToOrg($Org.Id, $AddUserIds)
            }
        }

        # 組織名変更
        if ($PSBoundParameters.NewOrganizationName) {
            if (Get-GrnOrganization $NewOrganizationName -NoDetail -URL $URL -Credential $Credential -ea SilentlyContinue -wa SilentlyContinue) {
                Write-Error ('{0} という組織が既に存在するため、組織名を変更できません' -f $NewOrganizationName)
                # 本当は組織名は重複可能だけど、ややこしいことになるのでエラー扱いにする
            }
            else {
                $private:Ret = $admin.ModifyOrgInfo($Org.Id, $Org.Code, $NewOrganizationName)
                if ($Ret.org_name) {
                    $OrganizationName = $Ret.org_name   #組織名を更新
                }
            }
        }

        if ($PassThru) {
            Get-GrnOrganization $OrganizationName -URL $URL -Credential $Credential -ErrorAction Continue
        }
    }

    End {}
}