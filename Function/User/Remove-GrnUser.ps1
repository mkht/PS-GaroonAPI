<#
.Synopsis
   ガルーンのユーザを削除します
.DESCRIPTION
   指定したログイン名のガルーンユーザを削除します
.PARAMETER URL
    ガルーンのURL
    必ずトップページのURLを指定してください
    例: http://grnserver/cgi-bin/cgi/grn.cgi
.PARAMETER Credential
    ガルーンに接続するための資格情報
.PARAMETER LoginName
    削除するユーザのログイン名
.EXAMPLE
    Remove-GrnUser -LoginName 'miyamizu' -URL $URL -Credential $cred
    Example 1: ログイン名が'miyamizu'のユーザを削除します
#>

function Remove-GrnUser {
    [CmdletBinding()]
    Param
    (
        # ガルーンのURL
        [Parameter(Mandatory)]
        [string]$URL,

        # ガルーン管理者の資格情報
        [Parameter(Mandatory)]
        [pscredential]$Credential,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)] [string]$LoginName, # ログイン名
        [parameter(DontShow, ValueFromPipelineByPropertyName)][int]$Id = -1    # ユーザID（隠しパラメータ）
    )

    Begin {
        $admin = New-Object GaroonAdmin @($URL, $Credential) -ErrorAction Stop
    }
    Process {
        try {
            $Id = $admin.GetUserIdByLoginName($LoginName)
            if ($Id -le 0) {
                throw ("ユーザ`"$LoginName`"が見つかりません")
            }
            else {
                Write-Verbose ("ユーザ`"$LoginName`"を削除します")
                $Removed = $admin.RemoveUsersByIds($Id)
                if ($Id -eq $Removed[0]) {
                    Write-Verbose 'ユーザが削除されました'
                }
                else {
                    Write-Warning 'Unknown Error'
                }
            }
        }
        catch [Exception] {
            Write-Error -Exception $_.Exception
            return
        }
    }

    End {}
}
