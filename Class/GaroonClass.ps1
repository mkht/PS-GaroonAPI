using namespace System.Xml
using namespace System.Security

# Garoon APIのラッパークラス群
# https://cybozudev.zendesk.com/hc/ja/categories/200157760-Garoon-API

# 基底クラス
class GaroonClass {
    # リクエストXMLのテンプレート
    [string]$RequestBase = @'
<?xml version="1.0" encoding="UTF-8"?>
<soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
  <soap:Header>
    <Action>
      {0}
    </Action>
    <Security>
      <UsernameToken>
        <Username>{2}</Username>
        <Password>{3}</Password>
      </UsernameToken>
    </Security>
    <Timestamp>
      <Created>{4}</Created>
      <Expires>{5}</Expires>
    </Timestamp>
    <Locale>jp</Locale>
  </soap:Header>
  <soap:Body>
  <{0}>
    {1}
  </{0}>
  </soap:Body>
</soap:Envelope>
'@

    #[string]$RequestXml

    [System.Management.Automation.PSCredential] $Credential # ガルーンログインアカウント情報
    [ValidatePattern('^https?://(.+(grn|index)(\.exe|\.cgi|\.csp)\??|.+\.cybozu\.com/g/)$')]
    [string] $GrnURL    # ガルーンのURL
    [ValidateSet('.exe', '.csp', '.cgi', '')]
    [string] $Ext   # [未使用]拡張子
    [string] $ApiSuffix
    [string] $RequestURI    # 実際にリクエストするURL
    Hidden [XmlDocument] $XmlDoc

    <# ---- コンストラクタ ---- #>
    GaroonClass() {
        $this.XmlDoc = New-Object XmlDocument
        $this.XmlDoc.PreserveWhitespace = $true
    }

    GaroonClass([string]$URL) {
        $this.XmlDoc = New-Object XmlDocument
        $this.XmlDoc.PreserveWhitespace = $true

        $this.GrnURL = $URL

        # for Garoon on cybozu
        if ($tURL = $([regex]::Match($URL, '.+cybozu.com/g'))[0].Value) {
            $this.RequestURI = $tURL + $this.ApiSuffix
        }
        # for Package version
        else {
            $this.RequestURI = $this.GrnURL + $this.ApiSuffix
        }
    }

    GaroonClass([string]$URL, [PSCredential] $Credential) {
        $this.XmlDoc = New-Object XmlDocument
        $this.XmlDoc.PreserveWhitespace = $true

        $this.GrnURL = $URL
        $this.Credential = $Credential

        # for Garoon on cybozu
        if ($tURL = $([regex]::Match($URL, '.+cybozu.com/g'))[0].Value) {
            $this.RequestURI = $tURL + $this.ApiSuffix
        }
        # for Package version
        else {
            $this.RequestURI = $this.GrnURL + $this.ApiSuffix
        }
    }

    static [xml]InvokeSOAPRequest ([Xml]$SOAPRequest, [String] $URL) {
        Write-Verbose "Sending SOAP Request To Server: $URL"
        $soapWebRequest = [System.Net.WebRequest]::Create($URL)
        $soapWebRequest.ContentType = 'text/xml;charset="utf-8"'
        $soapWebRequest.Accept = 'text/xml'
        $soapWebRequest.Method = 'POST'

        Write-Verbose 'Initiating Send.'
        $requestStream = $soapWebRequest.GetRequestStream()
        $SOAPRequest.Save($requestStream)
        $requestStream.Close()

        Write-Verbose 'Send Complete, Waiting For Response.'
        $resp = $soapWebRequest.GetResponse()
        $responseStream = $resp.GetResponseStream()
        $soapReader = [System.IO.StreamReader]($responseStream)
        $ReturnXml = [Xml] $soapReader.ReadToEnd()
        $responseStream.Close()

        Write-Verbose 'Response Received.'
        return $ReturnXml
    }

    # 引数とテンプレートからリクエスト用XMLを作る
    # @Private
    [string]CreateRequestXml([string]$Action, [string]$ParamterBody, [DateTime]$CreateTime) {
        $CreateUtcTime = [System.TimeZoneInfo]::ConvertTimeToUtc($CreateTime)   # GaroonAPIのTimestampはUTCのみ
        $ExpireUtcTime = $CreateUtcTime.AddHours(1) #期限は適当に1時間
        if ($this.Credential) {
            return ($this.RequestBase -f $Action, $ParamterBody, $this.Credential.UserName, $this.Credential.GetNetworkCredential().Password, $CreateUtcTime.ToString('yyyy-MM-ddTHH:mm:ssK'), $ExpireUtcTime.ToString('yyyy-MM-ddTHH:mm:ssK'))
        }
        else {
            return ($this.RequestBase -f $Action, $ParamterBody, '', '', $CreateUtcTime.ToString('yyyy-MM-ddTHH:mm:ssK'), $ExpireUtcTime.ToString('yyyy-MM-ddTHH:mm:ssK'))
        }
    }

    # リクエストを投げる
    # @Private
    [xml]Request([string]$RequestXml) {
        return $this.TestResponseXml([GaroonClass]::InvokeSOAPRequest($RequestXml, $this.RequestURI))
    }

    # 任意のAPIを叩く。 メソッドとして用意してないAPIを叩く必要があるとき用
    # @Public
    [Object[]]InvokeAnyApi($Action, $RequestBody) {
        $ResponseXml = $this.Request($this.CreateRequestXml($Action, $RequestBody, (Get-Date)))
        return $ResponseXml.Envelope.Body
    }

    # レスポンスがエラー情報を含んでいる場合はエラーを出す。正常なら入力XMLをそのまま返却
    # @Private
    [xml]TestResponseXml([xml]$ResponseXml) {
        if (-not $ResponseXml) {
            $msg = 'No Response returned.'
            # Write-Warning $msg
            throw $msg
        }
        if ($Fault = $ResponseXml.Envelope.Body.Fault) {
            $msg = ('[ERROR][{0}] {1}' -f $Fault.Detail.Code, $Fault.Reason.Text)
            # Write-Warning $msg
            throw $msg
        }
        return $ResponseXml
    }
}
