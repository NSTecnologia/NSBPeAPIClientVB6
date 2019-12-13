Attribute VB_Name = "BPeAPI"
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'activate Microsoft XML, v6.0 in references
Public responseText As String

'Atributo privado da classe
Private Const tempoResposta = 500
Private Const token = "SEU_TOKEN"

Function enviaConteudoParaAPI(conteudo As String, url As String, tpConteudo As String) As String
On Error GoTo SAI
    Dim contentType As String
    
    If (tpConteudo = "txt") Then
        contentType = "text/plain;charset=utf-8"
    ElseIf (tpConteudo = "xml") Then
        contentType = "application/xml;charset=utf-8"
    Else
        contentType = "application/json;charset=utf-8"
    End If
    
    Dim obj As MSXML2.ServerXMLHTTP60
    Set obj = New MSXML2.ServerXMLHTTP60
    obj.Open "POST", url
    obj.setRequestHeader "Content-Type", contentType
    If Trim(token) <> "" Then
        obj.setRequestHeader "X-AUTH-TOKEN", token
    End If
    obj.send conteudo
    Dim resposta As String
    resposta = obj.responseText
    
    Select Case obj.status
        'Se o token não for enviado ou for inválido
        Case 401
            MsgBox ("Token não enviado ou inválido")
        'Se o token informado for inválido 403
        Case 403
            MsgBox ("Token sem permissão")
    End Select
    
    enviaConteudoParaAPI = resposta
    Exit Function
SAI:
  enviaConteudoParaAPI = "{" & """status"":""" & Err.Number & """," & """motivo"":""" & Err.Description & """" & "}"
End Function

'Emitir BP-e Síncrono
Public Function emitirBPeSincrono(conteudo As String, tpConteudo As String, CNPJ As String, tpDown As String, tpAmb As String, caminho As String, exibeNaTela As Boolean) As String
    Dim retorno As String
    Dim resposta As String
    Dim statusEnvio As String
    Dim statusConsulta As String
    Dim statusDownload As String
    Dim motivo As String
    Dim erros As String
    Dim nsNRec As String
    Dim chBPe As String
    Dim cStat As String
    Dim nProt As String

    status = ""
    motivo = ""
    erros = ""
    nsNRec = ""
    chBPe = ""
    cStat = ""
    nProt = ""

    gravaLinhaLog ("[EMISSAO_SINCRONA_INICIO]")
    
    resposta = emitirBPe(conteudo, tpConteudo)
    statusEnvio = LerDadosJSON(resposta, "status", "", "")

    If (statusEnvio = "200") Or (statusEnvio = "-6") Then
        nsNRec = LerDadosJSON(resposta, "nsNRec", "", "")

        Sleep (tempoResposta)

        resposta = consultarStatusProcessamento(CNPJ, nsNRec, tpAmb)
        statusConsulta = LerDadosJSON(resposta, "status", "", "")

        If (statusConsulta = "200") Then
            cStat = LerDadosJSON(resposta, "cStat", "", "")

            If (cStat = "100") Then
                chBPe = LerDadosJSON(resposta, "chBPe", "", "")
                nProt = LerDadosJSON(resposta, "nProt", "", "")
                motivo = LerDadosJSON(resposta, "xMotivo", "", "")

                resposta = downloadBPeESalvar(chBPe, tpAmb, tpDown, caminho, exibeNaTela)
                statusDownload = LerDadosJSON(resposta, "status", "", "")

                If (statusDownload <> "200") Then
                    motivo = LerDadosJSON(resposta, "motivo", "", "")
                End If
            ElseIf (statusConsulta = "-2") Then
                erros = Split(resposta, """erro"":""")
                erros = LerDadosJSON(resposta, "erro", "", "")
        
                motivo = LerDadosJSON(erros, "xMotivo", "", "")
        
                cStat = LerDadosJSON(erros, "cStat", "", "")
            Else
                motivo = LerDadosJSON(resposta, "xMotivo", "", "")
            End If
        Else
            motivo = LerDadosJSON(resposta, "motivo", "", "")
        End If
   
    ElseIf (statusEnvio = "-4") Or (statusEnvio = "-2") Then

        motivo = LerDadosJSON(resposta, "motivo", "", "")
        erros = LerDadosJSON(resposta, "erros", "", "")

    ElseIf (statusEnvio = "-5") Then
        'Lê o objeto erro
        erros = Split(resposta, """erro"":""")
        erros = Left(erros, (Len(erros) - 1))
        erros = LerDadosJSON(erros, "erro", "", "")
        
        motivo = LerDadosJSON(erros, "xMotivo", "", "")
        
        cStat = LerDadosJSON(erros, "cStat", "", "")
    Else
    
        motivo = LerDadosJSON(resposta, "motivo", "", "")
        
    End If
    
    retorno = "{"
    retorno = retorno & """statusEnvio"":""" & statusEnvio & ""","
    retorno = retorno & """statusConsulta"":""" & statusConsulta & ""","
    retorno = retorno & """statusDownload"":""" & statusDownload & ""","
    retorno = retorno & """cStat"":""" & cStat & ""","
    retorno = retorno & """chBPe"":""" & chBPe & ""","
    retorno = retorno & """nProt"":""" & nProt & ""","
    retorno = retorno & """motivo"":""" & motivo & ""","
    retorno = retorno & """nsNRec"":""" & nsNRec & ""","
    retorno = retorno & """erros"":""" & erros & """"
    retorno = retorno & "}"
    
    'Grava dados de retorno
    gravaLinhaLog ("[JSON_RETORNO]")
    gravaLinhaLog (retorno)
    gravaLinhaLog ("[EMISSAO_SINCRONA_FIM]")

    emitirBPeSincrono = retorno
End Function


'Emitir BP-e
Public Function emitirBPe(conteudo As String, tpConteudo As String) As String
    Dim url As String
    Dim resposta As String
    
    url = "https://bpe.ns.eti.br/v1/bpe/issue"

    gravaLinhaLog ("[ENVIO_DADOS]")
    gravaLinhaLog (conteudo)
    
    resposta = enviaConteudoParaAPI(conteudo, url, tpConteudo)
    
    gravaLinhaLog ("[ENVIO_RESPOSTA]")
    gravaLinhaLog (resposta)

    emitirBPe = resposta
End Function

'Consultar Status de Processamento
Public Function consultarStatusProcessamento(CNPJ As String, nsNRec As String, tpAmb As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    json = "{"
    json = json & """CNPJ"":""" & CNPJ & ""","
    json = json & """nsNRec"":""" & nsNRec & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"

    url = "https://bpe.ns.eti.br/v1/bpe/issue/status"
    
    gravaLinhaLog ("[CONSULTA_DADOS]")
    gravaLinhaLog (json)
    
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[CONSULTA_RESPOSTA]")
    gravaLinhaLog (resposta)

    consultarStatusProcessamento = resposta
End Function

'Download do BP-e
Public Function downloadBPe(chBPe As String, tpDown As String, tpAmb As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim status As String

    json = "{"
    json = json & """chBPe"":""" & chBPe & ""","
    json = json & """tpDown"":""" & tpDown & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"

    url = "https://bpe.ns.eti.br/v1/bpe/get"

    gravaLinhaLog ("[DOWNLOAD_BPE_DADOS]")
    gravaLinhaLog (json)
    
    resposta = enviaConteudoParaAPI(json, url, "json")
    status = LerDadosJSON(resposta, "status", "", "")
    
    If (status <> "200") Then
        gravaLinhaLog ("[DOWNLOAD_BPE_RESPOSTA]")
        gravaLinhaLog (resposta)
    Else
        gravaLinhaLog ("[DOWNLOAD_BPE_RESPOSTA]")
        gravaLinhaLog (status)
    End If

    downloadBPe = resposta
End Function

'Download do BP-e e Salvar
Public Function downloadBPeESalvar(chBPe As String, tpAmb As String, tpDown As String, caminho As String, exibeNaTela As Boolean) As String
    Dim xml As String
    Dim json As String
    Dim pdf As String
    Dim status As String
    Dim resposta As String
    
    resposta = downloadBPe(chBPe, tpDown, tpAmb)
    status = LerDadosJSON(resposta, "status", "", "")

    If status = "200" Then
    
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
    
        'Checa se deve baixar XML
        If InStr(1, tpDown, "X") Then
            xml = LerDadosJSON(resposta, "xml", "", "")
            Call salvarXML(xml, caminho, chBPe, "", "")
        End If
        'Checa se deve baixar JSON
        If InStr(1, tpDown, "J") Then
            Dim conteudoJSON() As String
            'Separa o JSON da BPe
            conteudoJSON = Split(resposta, """BPeProc"":{")
            json = "{""BPeProc"":{" & conteudoJSON(1)
            Call salvarJSON(json, caminho, chBPe, "", "")
        End If
        'Checa se deve baixar PDF
        If InStr(1, tpDown, "P") Then
            pdf = LerDadosJSON(resposta, "pdf", "", "")
            Call salvarPDF(pdf, caminho, chBPe, "", "")
            
            If exibeNaTela Then
                'Abrindo o PDF gerado acima
                ShellExecute 0, "open", caminho & chBPe & "-procBPe.pdf", "", "", vbNormalFocus
            End If
        End If
    Else
        MsgBox ("Ocorreu um erro, veja o Retorno da API para mais informações")
    End If

    downloadBPeESalvar = resposta
End Function

'Download do Evento do BP-e
Public Function downloadEventoBPe(chBPe As String, tpAmb As String, tpDown As String, tpEvento As String, nSeqEvento As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    json = "{"
    json = json & """chBPe"":""" & chBPe & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & """tpDown"":""" & tpDown & ""","
    json = json & """tpEvento"":""" & tpEvento & ""","
    json = json & """nSeqEvento"":""" & nSeqEvento & """"
    json = json & "}"

    url = "https://bpe.ns.eti.br/v1/bpe/get/event"
    
    gravaLinhaLog ("[DOWNLOAD_EVENTO_DADOS]")
    gravaLinhaLog (json)
    
    resposta = enviaConteudoParaAPI(json, url, "json")
    status = LerDadosJSON(resposta, "status", "", "")
    
    If (status <> "200") Then
        gravaLinhaLog ("[DOWNLOAD_EVENTO_RESPOSTA]")
        gravaLinhaLog (resposta)
    Else
        gravaLinhaLog ("[DOWNLOAD_EVENTO_RESPOSTA]")
        gravaLinhaLog (status)
    End If

    downloadEventoBPe = resposta
End Function

'Download do Evento do BP-e e Salvar
Public Function downloadEventoBPeESalvar(chBPe As String, tpAmb As String, tpDown As String, tpEvento As String, nSeqEvento As String, caminho As String, exibeNaTela As Boolean) As String
    Dim baixarXML As Boolean
    Dim baixarPDF As Boolean
    Dim baixarJSON As Boolean
    Dim xml As String
    Dim json As String
    Dim pdf As String
    Dim status As String
    Dim resposta As String
    Dim tpEventoSalvar As String

    resposta = downloadEventoBPe(chBPe, tpAmb, tpDown, tpEvento, nSeqEvento)
    status = LerDadosJSON(resposta, "status", "", "")

    If status = "200" Then
        
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
        
        If (UCase(tpEvento) = "CANC") Then
          tpEventoSalvar = "110111"
        Else
          tpEventoSalvar = "110115"
        End If
        
        'Checa se deve baixar XML
        If InStr(1, tpDown, "X") Then
            xml = LerDadosJSON(resposta, "xml", "", "")
            Call salvarXML(xml, caminho, chBPe, tpEventoSalvar, nSeqEvento)
        End If
        'Checa se deve baixar JSON
        If InStr(1, tpDown, "J") Then
            json = LerDadosJSON(resposta, "json", "", "")
            Call salvarJSON(json, caminho, chBPe, tpEventoSalvar, nSeqEvento)
        End If
        'Checa se deve baixar PDF
        If InStr(1, tpDown, "P") Then
            pdf = LerDadosJSON(resposta, "pdf", "", "")
            Call salvarPDF(pdf, caminho, chBPe, tpEventoSalvar, nSeqEvento)
            
            If exibeNaTela Then
                'Abrindo o PDF gerado acima
                ShellExecute 0, "open", caminho & tpEventoSalvar & chBPe & nSeqEvento & "-procEvenBPe.pdf", "", "", vbNormalFocus
            End If
        End If
    Else
        MsgBox ("Ocorreu um erro, veja o Retorno da API para mais informações")
    End If

    downloadEventoBPeESalvar = resposta
End Function

'Realizar o cancelamento do BP-e
Public Function cancelarBPe(chBPe As String, tpAmb As String, dhEvento As String, nProt As String, xJust As String, tpDown As String, caminho As String, exibeNaTela As Boolean) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim status As String
    Dim respostaDownload As String

    json = "{"
    json = json & """chBPe"":""" & chBPe & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & """dhEvento"":""" & dhEvento & ""","
    json = json & """nProt"":""" & nProt & ""","
    json = json & """xJust"":""" & xJust & """"
    json = json & "}"
    
    url = "https://bpe.ns.eti.br/v1/bpe/cancel"
    
    gravaLinhaLog ("[CANCELAMENTO_DADOS]")
    gravaLinhaLog (json)
    
    resposta = enviaConteudoParaAPI(token, json, url, "json")

    gravaLinhaLog ("[CANCELAMENTO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    status = LerDadosJSON(resposta, "status", "", "")
    
    If (status = "200") Then
        respostaDownload = downloadEventoBPeESalvar(chBPe, tpAmb, tpDown, "CANC", "1", caminho, exibeNaTela)
        status = LerDadosJSON(respostaDownload, "status", "", "")
        
        If (status <> "200") Then
            MsgBox ("Ocorreu um erro ao fazer o download. Verifique os logs.")
        End If
    End If
    
    cancelarBPe = resposta
End Function

'Realizar o evento de nao embarque de um CT-e
Public Function naoEmbarqueBPe(chBPe As String, tpAmb As String, dhEvento As String, nProt As String, xJust As String, tpDown As String, caminho As String, exibeNaTela As Boolean) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim status As String
    Dim respostaDownload As String
    
    
    json = "{"
    json = json & """chBPe"":""" & chBPe & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & """dhEvento"":""" & dhEvento & ""","
    json = json & """nProt"":""" & nProt & ""","
    json = json & """xJust"":""" & xJust & """"
    json = json & "}"
    
    url = "https://bpe.ns.eti.br/v1/bpe/naoemb"
    
    gravaLinhaLog ("[NAO_EMB_DADOS]")
    gravaLinhaLog (json)
  
    resposta = enviaConteudoParaAPI(json, url, "json")

    gravaLinhaLog ("[[NAO_EMB_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    status = LerDadosJSON(resposta, "status", "", "")
    
    If (status = "200") Then
        respostaDownload = downloadEventoBPeESalvar(chBPe, tpAmb, tpDown, "NAO_EMB", "1", caminho, exibeNaTela)
        status = LerDadosJSON(respostaDownload, "status", "", "")
        
        If (status <> "200") Then
            MsgBox ("Ocorreu um erro ao fazer o download. Verifique os logs.")
        End If
    End If
    
    corrigirBPe = resposta
End Function

'Esta função realiza a consulta de situação de um BP-e
Public Function consultarSituacao(licencaCnpj As String, chBPe As String, tpAmb As String, versao As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """licencaCnpj"":""" & licencaCnpj & ""","
    json = json & """chBPe"":""" & chBPe & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"

    url = "'https://bpe.ns.eti.br/v1/bpe/status"
    
    gravaLinhaLog ("[CONSULTA_SITUACAO_DADOS]")
    gravaLinhaLog (json)

    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[CONSULTA_SITUACAO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    consultarSituacao = resposta
End Function

'Salvar XML
Public Sub salvarXML(xml As String, caminho As String, chBPe As String, tpEvento As String, nSeqEvento As String)
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String

    'Seta o caminho para o arquivo XML
    If (nSeqEvento = "") Then
        localParaSalvar = caminho & chBPe & nSeqEvento & "-procBPe.xml"
    Else
        localParaSalvar = caminho & tpEvento & chBPe & nSeqEvento & "-procEvenBPe.xml"
    End If

    'Remove as contrabarras
    conteudoSalvar = Replace(xml, "\", "")

    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open
    fsT.WriteText conteudoSalvar
    fsT.SaveToFile localParaSalvar
End Sub

'Salvar JSON
Public Sub salvarJSON(json As String, caminho As String, chBPe As String, tpEvento As String, nSeqEvento As String)
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String

    'Seta o caminho para o arquivo JSON
    If (nSeqEvento = "") Then
        localParaSalvar = caminho & chBPe & nSeqEvento & "-procBPe.json"
    Else
        localParaSalvar = caminho & tpEvento & chBPe & nSeqEvento & "-procEvenBPe.json"
    End If

    conteudoSalvar = json

    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open
    fsT.WriteText conteudoSalvar
    fsT.SaveToFile localParaSalvar
End Sub

'Salvar PDF
Public Function salvarPDF(pdf As String, caminho As String, chBPe As String, tpEvento As String, nSeqEvento As String) As Boolean
On Error GoTo SAI
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String

    'Seta o caminho para o arquivo PDF
    If (nSeqEvento = "") Then
        localParaSalvar = caminho & chBPe & nSeqEvento & "-procBPe.pdf"
    Else
        localParaSalvar = caminho & tpEvento & chBPe & nSeqEvento & "-procEvenBPe.pdf"
    End If

    Dim fnum
    fnum = FreeFile
    Open localParaSalvar For Binary As #fnum
    Put #fnum, 1, Base64Decode(pdf)
    Close fnum
    Exit Function
SAI:
    MsgBox (Err.Number & " - " & Err.Description), vbCritical
End Function

'activate microsoft script control 1.0 in references
Public Function LerDadosJSON(sJsonString As String, Key1 As String, Key2 As String, key3 As String) As String
On Error GoTo err_handler
    Dim oScriptEngine As ScriptControl
    Set oScriptEngine = New ScriptControl
    oScriptEngine.Language = "JScript"
    Dim objJSON As Object
    Set objJSON = oScriptEngine.Eval("(" + sJsonString + ")")
    If Key1 <> "" And Key2 <> "" And key3 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, Key1, VbGet), Key2, VbGet), key3, VbGet)
    ElseIf Key1 <> "" And Key2 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(objJSON, Key1, VbGet), Key2, VbGet)
    ElseIf Key1 <> "" Then
        LerDadosJSON = VBA.CallByName(objJSON, Key1, VbGet)
    End If
Err_Exit:
    Exit Function
err_handler:
    LerDadosJSON = "Error: " & Err.Description
    Resume Err_Exit
End Function

Public Function LerDadosXML(sXml As String, Key1 As String, Key2 As String) As String
    On Error Resume Next
    LerDadosXML = ""
    
    Set xml = New DOMDocument60
    xml.async = False
    
    If xml.loadXML(sXml) Then
        'Tentar pegar o strCampoXML
        Set objNodeList = xml.getElementsByTagName(Key1 & "//" & Key2)
        Set objNode = objNodeList.nextNode
        
        Dim valor As String
        valor = objNode.Text
        
        If Len(Trim(valor)) > 0 Then 'CONSEGUI LER O XML NODE
            LerDadosXML = valor
        End If
        Else
        MsgBox "Não foi possível ler o conteúdo do XML do BPe especificado para leitura.", vbCritical, "ERRO"
    End If
End Function

'Função genérica para gravação de log
Public Sub gravaLinhaLog(conteudoSalvar As String)
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim localParaSalvar As String
    Dim data As String
    
    'Diretório para salvar os logs
    localParaSalvar = App.Path & "\log\"
    
    'Checa se existe o caminho passado para salvar os arquivos
    If Dir(localParaSalvar, vbDirectory) = "" Then
        MkDir (localParaSalvar)
    End If
    
    'Pega data atual
    data = Format(Date, "yyyyMMdd")
    
    'Diretório + nome do arquivo para salvar os logs
    localParaSalvar = App.Path & "\log\" & data & ".txt"
    
    'Pega data e hora atual
    data = DateTime.Now
    
    Dim fnum
    fnum = FreeFile
    Open localParaSalvar For Append Shared As #fnum
    Print #fnum, data & " - " & conteudoSalvar
    Close fnum
End Sub

