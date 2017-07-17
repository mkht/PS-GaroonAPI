
function New-GrnMailAccount
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
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)] [string]$Password,   # パスワード
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)] [string]$Email,   # メールアドレス
        [Parameter(ValueFromPipelineByPropertyName)] [string]$AccountCode,    # アカウントコード
        [Parameter(ValueFromPipelineByPropertyName)] [string]$AccountName,    # アカウント名
        [Parameter(ValueFromPipelineByPropertyName)] [bool]$LeaveServerMail = $false,   # メールサーバーにメールを残す
        [Parameter(ValueFromPipelineByPropertyName)] [string]$Invalid = $false,    # アカウント停止
        [parameter(DontShow,ValueFromPipelineByPropertyName)][string]$Id    # ユーザID（隠しパラメータ）
    )

    Begin
    {
        $base = New-Object GaroonBase @($URL,$Credential) -ErrorAction Stop
        $mail = New-Object GaroonMail @($URL,$Credential) -ErrorAction Stop
    }
    Process
    {

        if(-not $AccountCode){
            $AccountCode = $Email
        }
        if(-not $AccountName){
            $AccountName = $Email
        }

        try{
            $user = $base.GetUsersByLoginName($LoginName)
        }
        catch [Exception]{
            # GRNERR以外のエラー(404など)の場合は中断
            if($_.Exception -is [System.Net.WebException]){
                Write-Error -Exception $_.Exception
                return
            }
        }
        if(-not ($Id = $user.key)){
            $msg = ("ユーザ'{0}'が存在しません" -f $LoginName)
            Write-Error $msg
            return
        }

        $MailParams = @{
            AccountId = 1
            UserId = $Id
            UserAccountCode = $UserAccountCode
            UserAccountName = $UserAccountName
            MailServerId = 1
            Email = $Email
            AccountName = $UserAccountName
            Password = $Password
            LeaveServerMail = $LeaveServerMail
            Deactivate = $Invalid
        }

        $MailAccountInfo = New-Object MailUserAccount $MailParams

        try{
            $result = $mail.CreateUserAccount($MailAccountInfo)
        }
        catch [Exception]{
            # GRNERR以外のエラー(404など)の場合は中断
            if($_.Exception -is [System.Net.WebException]){
                Write-Error -Exception $_.Exception
                return
            }
            elseif ($_.Exception.Message -match 'GRN_MAIL_24105') {
                Write-Error ("このユーザは既にメールアカウントが設定されています")
                return
            }
            else{
                Write-Error -Exception $_.Exception
                return
            }
        }

        if($PassThru){
            $result
        }
    }
    End {}
}