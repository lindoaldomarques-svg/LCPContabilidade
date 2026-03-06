Attribute VB_Name = "BBExtratos"
Option Explicit

' ================================================================
' Módulo VBA para consumir a API de Extratos do Banco do Brasil
' e gravar o resultado em uma planilha do Excel.
'
' Pré-requisitos comuns da API BB:
' 1) Aplicação cadastrada no Developer BB.
' 2) OAuth2 (Client Credentials) habilitado.
' 3) Certificado cliente (mTLS), quando exigido pelo convênio.
' 4) Escopos/permissões para endpoint de extratos.
'
' Observação importante:
' - Os endpoints abaixo são exemplos de URL. Valide no portal do BB
'   os caminhos exatos do seu produto/versão e ambiente.
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
' Exemplo: "CURRENT_USER\\MY\\THUMBPRINT_DO_CERTIFICADO"
Private Const CLIENT_CERT_LOCATION As String = ""

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
    token = ObterTokenBB()
    If Len(token) = 0 Then
        MsgBox "Não foi possível obter token OAuth2.", vbCritical
        Exit Sub
    End If

    Dim json As String
    json = BuscarExtratoBB(token, dtInicio, dtFim)
    If Len(json) = 0 Then
        MsgBox "Resposta vazia da API de extratos.", vbExclamation
        Exit Sub
    End If

    GravarJsonEmPlanilha json, "Extrato_BB_JSON"

    MsgBox "Extrato baixado com sucesso.", vbInformation
    Exit Sub

TrataErro:
    MsgBox "Erro ao baixar extrato BB: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ====== OAUTH2 ======
Private Function ObterTokenBB() As String
    On Error GoTo TrataErro

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", BB_AUTH_URL, False

    If Len(Trim$(CLIENT_CERT_LOCATION)) > 0 Then
        http.SetClientCertificate CLIENT_CERT_LOCATION
    End If

    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.SetRequestHeader "Authorization", "Basic " & Base64Encode(CLIENT_ID & ":" & CLIENT_SECRET)

    Dim body As String
    body = "grant_type=client_credentials"

    If Len(Trim$(OAUTH_SCOPE)) > 0 Then
        body = body & "&scope=" & UrlEncode(OAUTH_SCOPE)
    End If

    http.Send body

    If http.Status < 200 Or http.Status >= 300 Then
        Debug.Print "Falha OAuth2: "; http.Status; " - "; http.ResponseText
        ObterTokenBB = ""
        Exit Function
    End If

    ObterTokenBB = JsonGetString(http.ResponseText, "access_token")
    Exit Function

TrataErro:
    Debug.Print "Erro ObterTokenBB: "; Err.Number; " - "; Err.Description
    ObterTokenBB = ""
End Function

' ====== EXTRATOS ======
Private Function BuscarExtratoBB(ByVal accessToken As String, ByVal dataInicio As String, ByVal dataFim As String) As String
    On Error GoTo TrataErro

    Dim url As String
    url = Replace(BB_EXTRATO_URL, "{agencia}", UrlEncode(AGENCIA))
    url = Replace(url, "{conta}", UrlEncode(CONTA))

    url = url & "?gw-dev-app-key=" & UrlEncode(APP_KEY)
    url = url & "&dataInicioSolicitacao=" & UrlEncode(dataInicio)
    url = url & "&dataFimSolicitacao=" & UrlEncode(dataFim)

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "GET", url, False

    If Len(Trim$(CLIENT_CERT_LOCATION)) > 0 Then
        http.SetClientCertificate CLIENT_CERT_LOCATION
    End If

    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Accept", "application/json"

    http.Send

    If http.Status < 200 Or http.Status >= 300 Then
        Debug.Print "Falha Extrato: "; http.Status; " - "; http.ResponseText
        BuscarExtratoBB = ""
        Exit Function
    End If

    BuscarExtratoBB = http.ResponseText
    Exit Function

TrataErro:
    Debug.Print "Erro BuscarExtratoBB: "; Err.Number; " - "; Err.Description
    BuscarExtratoBB = ""
End Function

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

' Parser mínimo para extração de string de uma chave JSON simples.
' Em produção, prefira um parser JSON completo (ex.: VBA-JSON).
Private Function JsonGetString(ByVal json As String, ByVal key As String) As String
    Dim token As String
    token = """" & key & """"

    Dim p As Long
    p = InStr(1, json, token, vbTextCompare)
    If p = 0 Then Exit Function

    p = InStr(p + Len(token), json, ":", vbTextCompare)
    If p = 0 Then Exit Function

    p = InStr(p + 1, json, """", vbTextCompare)
    If p = 0 Then Exit Function

    Dim q As Long
    q = InStr(p + 1, json, """", vbTextCompare)
    If q = 0 Then Exit Function

    JsonGetString = Mid$(json, p + 1, q - p - 1)
End Function
