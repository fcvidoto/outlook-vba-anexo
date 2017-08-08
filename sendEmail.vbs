'faz o envio de email pela instancia do 'Outlook'
Public Sub enviarEmailCN(ByRef cnList As Object, ByRef emailList As Object)
  Dim i As Integer
  Dim CriandoObjetos As CriandoObjetos
  Set CriandoObjetos = New CriandoObjetos
  Dim nameFormatFile As String
  Dim nomeArquivo As String
  Dim nomeArquivoNivelCTC As String
  Dim listaSenders As String
  '----------------------------
  'pega os valores do senders '<--
  For i = 0 To emailList.ListCount - 1
    If emailList.Selected(i) Then
      listaSenders = listaSenders & emailList.List(i) & "; "
    End If
  Next i
  '-----------------------------
  'verifica se foram selecionados valores nas 2 listbox
  If IsNull(cnList.Value) Or listaSenders = "" Then
      MsgBox "Selecione uma 'CN' e algum email da lista de emails", vbCritical
      Exit Sub
  End If
  '-----------------------------
  'verifica se o arquivo da pasta do anexo existe
  Call Abre_Banco 'Carrega os campos do formulario de acordo com a selecao na lista
  sSQL = "select * from tabGerarPEC where codigoCN=" & cnList.Value & " order by codigoCN ASC"
  '------------------------------
  rstlista.Open sSQL, dbADO, adOpenKeyset, adLockOptimistic
  nameFormatFile = Format(rstlista.Fields("codigoCN").Value, "00") & ". " & rstlista.Fields("CN").Value
  nomeArquivo = ThisWorkbook.path & "\PEC\" & "PEC - " & nameFormatFile & ".xlsb"
  nomeArquivoNivelCTC = ThisWorkbook.path & "\PEC\" & "PEC - " & nameFormatFile & " (Nível CTC).xlsb"
  '------------------------------
  If CriandoObjetos.ArquivoExiste(nomeArquivo) = False Or _
     CriandoObjetos.ArquivoExiste(nomeArquivoNivelCTC) = False Then
    MsgBox "Os arquivos da 'CN' selecionada: '" & rstlista.Fields("CN").Value & "' não existem na pasta: '" & ThisWorkbook.path & "\PEC\" & "'", vbCritical
    Exit Sub
  End If
  '-----------------------------
  Call enviaEmail(listaSenders, nomeArquivo, nomeArquivoNivelCTC, rstlista.Fields("CN").Value) 'envia o email com os 2 anexos a lista de senders montada pela variável // com os 2 anexos
  MsgBox "Email enviado ao 'Outlook' com sucesso", vbInformation
End Sub

'envia o email com os 2 anexos a lista de senders montada pela variável // com os 2 anexos
Public Sub enviaEmail(listaSenders As String, nomeArquivo As String, nomeArquivoNivelCTC As String, cn As String)
  Dim oOutlook As Object
  Dim oMSG As Object
  Dim oRecipient As Object
  Dim objOutlookAttach As Object
  Dim oNameSpace As Object
  Dim oExplorer As Object
  Dim oFolder As Object
  '-----------------------------
  Set oOutlook = CreateObject("Outlook.Application")
  Set oMSG = oOutlook.CreateItem(0) 'olMailItem = 0
  '-----------------------------
  Set oNameSpace = oOutlook.GetNamespace("MAPI")
  Set oFolder = oNameSpace.GetDefaultFolder(6) ' 6 = olFolderInbox
  Set oExplorer = oOutlook.Explorers.Add(oFolder, 0) ' 0 = olFolderDisplayNormal
  oExplorer.Activate
  '-----------------------------
  'monta a monta o email
  With oMSG
    Set oRecipient = .Recipients.Add(listaSenders) 'monta os senders
    '-----------------------------
    'anexa os arquivos do Excel
    .Attachments.Add nomeArquivo
    .Attachments.Add nomeArquivoNivelCTC
    '-----------------------------
    .To = listaSenders
    .Subject = "Arquivos da CN: '" & cn & "'"
    .Body = "Att:."
    .Save 'testes para salvar o email
    '.Display '<-- para testes
    .Send 'envia o email
  End With
End Sub
