' AbuseIPDBCheck.vbs

Option Explicit

' Define AbuseIPDB API key and IP address to check
Const apiKey = "ab5fb6b4d8828d26e898841a3f43d75117e5578da591a7056b430452e23c49cd415b325aa75f741e"
Const ip = "26.0.0.1"

' Create URL to AbuseIPDB API endpoint
Dim apiUrl
apiUrl = "https://api.abuseipdb.com/api/v2/check?ipAddress=" & ip

' Create HTTP request object
Dim objHTTP
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

' Open HTTP request with specified method and URL
objHTTP.Open "GET", apiUrl, False

' Set request headers
objHTTP.setRequestHeader "Key", apiKey
objHTTP.setRequestHeader "Accept", "application/json"

' Send HTTP request
objHTTP.Send

' Check if request was successful (status code 200)
If objHTTP.Status = 200 Then
    ' Output JSON response
    WScript.Echo objHTTP.responseText
Else
    ' Display error message if request fails
    WScript.Echo "Error: " & objHTTP.Status & " - " & objHTTP.statusText
End If

' Release HTTP request object
Set objHTTP = Nothing
