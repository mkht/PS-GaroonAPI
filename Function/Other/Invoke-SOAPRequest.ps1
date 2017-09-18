function Invoke-SOAPRequest 
{
    [CmdletBinding()]
    [OutputType([Xml])]
    Param(
        [Xml]    $SOAPRequest, 
        [String] $URL 
    )
    
    Write-Verbose "Sending SOAP Request To Server: $URL" 
    $soapWebRequest = [System.Net.WebRequest]::Create($URL)

    $soapWebRequest.ContentType = 'text/xml;charset="utf-8"'
    $soapWebRequest.Accept      = "text/xml"
    $soapWebRequest.Method      = "POST"
    
    $requestStream = $soapWebRequest.GetRequestStream() 
    $SOAPRequest.Save($requestStream) 
    $requestStream.Close() 
    
    Write-Verbose "Send Complete, Waiting For Response." 
    $resp = $soapWebRequest.GetResponse() 
    $responseStream = $resp.GetResponseStream() 
    $soapReader = [System.IO.StreamReader]($responseStream) 
    $ReturnXml = [Xml] $soapReader.ReadToEnd() 
    $responseStream.Close() 
    
    Write-Verbose "Response Received."

    return $ReturnXml 
}