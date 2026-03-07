Attribute VB_Name = "BBExtratos"
Option Explicit

' ================================================================
' Módulo VBA para consumir a API de Extratos do Banco do Brasil
' e gravar o resultado em uma planilha do Excel.
' ================================================================

' ====== CONFIGURAÇÃO ======
Private Const BB_AUTH_URL As String = "https://oauth.bb.com.br/oauth/token"
Private Const BB_EXTRATO_URL As String = "https://api.bb.com.br/extratos/v1/contas/{agencia}/{conta}/lancamentos"

Private Const CLIENT_ID As String = "SEU_CLIENT_ID"
Private Const CLIENT_SECRET As String = "SEU_CLIENT_SECRET"
Private Const APP_KEY As String = "SUA_APP_KEY"

' Formato recomendado para agência/conta sem máscara.
Private Const AGENCIA As String = "1234"
Private Const CONTA As String = "1234567"

' Escopo deve refletir o autorizado no BB Developer.
Private Const OAUTH_SCOPE As String = "extrato.read"

' Opcional para cenários mTLS em máquinas Windows.
' Exemplo: "CURRENT_USER\MY\THUMBPRINT_DO_CERTIFICADO"
Private Const CLIENT_CERT_LOCATION As String = ""

' Timeout padrão (ms)
Private Const HTTP_TIMEOUT_MS As Long = 60000

' ====== API PÚBLICA ======
Public Sub BaixarExtratoBB()
    On Error GoTo TrataErro

    Dim dtInicio As String
    Dim dtFim As String

    dtInicio = InputBox("Data inicial (AAAA-MM-DD):", "Extrato BB", Format(Date - 7, "yyyy-mm-dd"))
    If Len(Trim$(dtInicio)) = 0 Then Exit Sub

    dtFim = InputBox("Data final (AAAA-MM-DD):", "Extrato BB", Format(Date, "yyyy-mm-dd"))
    If Len(Trim$(dtFim)) = 0 Then Exit Sub

    If Not IsIsoDate(dtInicio) Or Not IsIsoDate(dtFim) Then
        MsgBox "Informe datas no formato AAAA-MM-DD.", vbExclamation
        Exit Sub
    End If

    Dim token As String
    Dim authErro As String

    token = ObterTokenBB(authErro)
    If Len(token) = 0 Then
        MsgBox "Não foi possível obter token OAuth2." & vbCrLf & vbCrLf & authErro, vbCritical
        Exit Sub
    End If

    Dim json As String
    Dim extratoErro As String

    json = BuscarExtratoBB(token, dtInicio, dtFim, extratoErro)
    If Len(json) = 0 Then
        MsgBox "Falha ao consultar extrato." & vbCrLf & vbCrLf & extratoErro, vbCritical
        Exit Sub
    End If

    GravarJsonEmPlanilha json, "Extrato_BB_JSON"

    MsgBox "Extrato baixado com sucesso.", vbInformation
    Exit Sub

TrataErro:
    MsgBox "Erro ao baixar extrato BB: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ====== OAUTH2 ======
Private Function ObterTokenBB(ByRef erroDetalhado As String) As String
    On Error GoTo TrataErro

    Dim tokenUrl As String
    tokenUrl = MontarUrlToken()

    Dim body As String
    body = "grant_type=client_credentials"
    If Len(Trim$(OAUTH_SCOPE)) > 0 Then
        body = body & "&scope=" & UrlEncode(OAUTH_SCOPE)
    End If

    ' 1ª tentativa: padrão OAuth2 (Authorization: Basic)
    Dim status As Long
    Dim resp As String
    status = ExecutarTokenRequest(tokenUrl, body, True, resp)

    ' 2ª tentativa (fallback): alguns ambientes aceitam client_id/client_secret no body
    If status < 200 Or status >= 300 Then
        Dim bodyFallback As String
        bodyFallback = body & _
            "&client_id=" & UrlEncode(CLIENT_ID) & _
            "&client_secret=" & UrlEncode(CLIENT_SECRET)

        status = ExecutarTokenRequest(tokenUrl, bodyFallback, False, resp)
    End If

    If status < 200 Or status >= 300 Then
        erroDetalhado = "HTTP " & CStr(status) & " no OAuth." & vbCrLf & _
                        "URL: " & tokenUrl & vbCrLf & _
                        "Resposta: " & LimitarTexto(resp, 600)
        ObterTokenBB = ""
        Exit Function
    End If

    ObterTokenBB = JsonGetString(resp, "access_token")

    If Len(ObterTokenBB) = 0 Then
        erroDetalhado = "OAuth retornou sucesso, porém sem access_token." & vbCrLf & _
                        "Resposta: " & LimitarTexto(resp, 600)
        Exit Function
    End If

    erroDetalhado = ""
    Exit Function

TrataErro:
    erroDetalhado = "Erro VBA no OAuth: " & Err.Number & " - " & Err.Description
    ObterTokenBB = ""
End Function

Private Function ExecutarTokenRequest(ByVal url As String, ByVal body As String, ByVal usarBasicAuth As Boolean, ByRef resposta As String) As Long
    On Error GoTo TrataErro

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ConfigurarHttp http

    http.Open "POST", url, False

    If Len(Trim$(CLIENT_CERT_LOCATION)) > 0 Then
        http.SetClientCertificate CLIENT_CERT_LOCATION
    End If

    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.SetRequestHeader "Accept", "application/json"

    If usarBasicAuth Then
        http.SetRequestHeader "Authorization", "Basic " & Base64Encode(CLIENT_ID & ":" & CLIENT_SECRET)
    End If

    http.Send body

    resposta = http.ResponseText
    ExecutarTokenRequest = http.Status
    Exit Function

TrataErro:
    resposta = "Erro de transporte: " & Err.Number & " - " & Err.Description
    ExecutarTokenRequest = 0
End Function

Private Function MontarUrlToken() As String
    If InStr(1, BB_AUTH_URL, "gw-dev-app-key=", vbTextCompare) > 0 Then
        MontarUrlToken = BB_AUTH_URL
        Exit Function
    End If

    If InStr(1, BB_AUTH_URL, "?", vbBinaryCompare) > 0 Then
        MontarUrlToken = BB_AUTH_URL & "&gw-dev-app-key=" & UrlEncode(APP_KEY)
    Else
        MontarUrlToken = BB_AUTH_URL & "?gw-dev-app-key=" & UrlEncode(APP_KEY)
    End If
End Function

' ====== EXTRATOS ======
Private Function BuscarExtratoBB(ByVal accessToken As String, ByVal dataInicio As String, ByVal dataFim As String, ByRef erroDetalhado As String) As String
    On Error GoTo TrataErro

    Dim url As String
    url = Replace(BB_EXTRATO_URL, "{agencia}", UrlEncode(AGENCIA))
    url = Replace(url, "{conta}", UrlEncode(CONTA))
    url = url & "?gw-dev-app-key=" & UrlEncode(APP_KEY)
    url = url & "&dataInicioSolicitacao=" & UrlEncode(dataInicio)
    url = url & "&dataFimSolicitacao=" & UrlEncode(dataFim)

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ConfigurarHttp http

    http.Open "GET", url, False

    If Len(Trim$(CLIENT_CERT_LOCATION)) > 0 Then
        http.SetClientCertificate CLIENT_CERT_LOCATION
    End If

    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Accept", "application/json"

    http.Send

    If http.Status < 200 Or http.Status >= 300 Then
        erroDetalhado = "HTTP " & CStr(http.Status) & " na API de extrato." & vbCrLf & _
                        "URL: " & url & vbCrLf & _
                        "Resposta: " & LimitarTexto(http.ResponseText, 600)
        BuscarExtratoBB = ""
        Exit Function
    End If

    erroDetalhado = ""
    BuscarExtratoBB = http.ResponseText
    Exit Function

TrataErro:
    erroDetalhado = "Erro VBA na API de extrato: " & Err.Number & " - " & Err.Description
    BuscarExtratoBB = ""
End Function

Private Sub ConfigurarHttp(ByVal http As Object)
    On Error Resume Next
    http.SetTimeouts HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS

    ' Força TLS 1.2 no WinHTTP (0x800)
    http.Option(9) = 2048
    http.Option(6) = True ' habilita redirects
    On Error GoTo 0
End Sub

' ====== SAÍDA ======
Private Sub GravarJsonEmPlanilha(ByVal json As String, ByVal nomeAba As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomeAba)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nomeAba
    End If

    ws.Cells.Clear
    ws.Range("A1").Value = "timestamp"
    ws.Range("B1").Value = "json"

    ws.Range("A2").Value = Now
    ws.Range("B2").Value = json

    ws.Columns("A:B").EntireColumn.AutoFit
End Sub

' ====== HELPERS ======
Private Function Base64Encode(ByVal plainText As String) As String
    Dim arrData() As Byte
    arrData = StrConv(plainText, vbFromUnicode)

    Dim dom As Object
    Dim node As Object

    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = dom.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = arrData

    Base64Encode = Replace(node.Text, vbLf, "")
End Function

Private Function UrlEncode(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    Dim ascCode As Integer
    Dim result As String

    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        ascCode = Asc(ch)

        Select Case ascCode
            Case 48 To 57, 65 To 90, 97 To 122
                result = result & ch
            Case 45, 46, 95, 126
                result = result & ch
            Case 32
                result = result & "%20"
            Case Else
                result = result & "%" & Right$("0" & Hex(ascCode), 2)
        End Select
    Next i

    UrlEncode = result
End Function

Private Function IsIsoDate(ByVal txt As String) As Boolean
    On Error GoTo Falha

    If Len(txt) <> 10 Then
        IsIsoDate = False
        Exit Function
    End If

    If Mid$(txt, 5, 1) <> "-" Or Mid$(txt, 8, 1) <> "-" Then
        IsIsoDate = False
        Exit Function
    End If

    Dim y As Integer, m As Integer, d As Integer
    y = CInt(Left$(txt, 4))
    m = CInt(Mid$(txt, 6, 2))
    d = CInt(Right$(txt, 2))

    Dim dt As Date
    dt = DateSerial(y, m, d)

    IsIsoDate = (Year(dt) = y And Month(dt) = m And Day(dt) = d)
    Exit Function

Falha:
    IsIsoDate = False
End Function

Private Function JsonGetString(ByVal json As String, ByVal key As String) As String
    On Error GoTo Fim

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    re.Global = False
    re.IgnoreCase = True
    re.MultiLine = True
    re.Pattern = """" & key & """" & "\s*:\s*""([^""]+)"""

    Dim matches As Object
    Set matches = re.Execute(json)

    If matches.Count > 0 Then
        JsonGetString = matches(0).SubMatches(0)
    End If

Fim:
End Function

Private Function LimitarTexto(ByVal texto As String, ByVal limite As Long) As String
    If Len(texto) <= limite Then
        LimitarTexto = texto
    Else
        LimitarTexto = Left$(texto, limite) & "..."
    End If
End Function
