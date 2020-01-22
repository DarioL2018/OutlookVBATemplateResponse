VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents oExpl As Explorer
Attribute oExpl.VB_VarHelpID = -1
Private WithEvents oItem As MailItem
Attribute oItem.VB_VarHelpID = -1

Private Sub Application_Startup()
   Set oExpl = Application.ActiveExplorer
   'bDiscardEvents = False
End Sub

Private Sub oExpl_SelectionChange()
   On Error Resume Next
   Set oItem = oExpl.Selection.Item(1)
End Sub

Private Sub oItem_Reply(ByVal Response As Object, Cancel As Boolean)
    
    'Set olReply = oItem.Reply
    Dim olInspector As Inspector
    Dim saludo As String
    
    Dim nombre As String
    Dim inicio As String
    'oItem.HTMLBody = "Hello, Thank you. " & vbCrLf & Response.HTMLBody
    'If Not olItem Is Nothing And olItem.Class = olMail Then
    'MsgBox "Fin"
    Set olInspector = Response.GetInspector
    If olInspector.EditorType = olEditorWord Then
        Set objWDoc = olInspector.WordEditor
        Set objWSel = objWDoc.Application.Selection
        
        saludo = getSaludo
        
        'Response.HTMLBody = "Estimad@, " & vbCrLf & saludo & vbCrLf & Response.HTMLBody
        Response.Display
        inicio = getCabecera(Response.Recipients)
        SendKeys inicio, True
        SendKeys "{ENTER}", True
        SendKeys saludo, True
        SendKeys "{ENTER}", True
        SendKeys "{ENTER}", True
        SendKeys "{ENTER}", True
        SendKeys "{ENTER}", True
        SendKeys "Saludos ", True
        SendKeys "{UP}", True
        SendKeys "{UP}", True
        
    End If
        
End Sub

Private Function getGeneroDominio(oExUser As Outlook.ExchangeUser) As String
    Dim genero As String
    If oExUser Is Nothing Then
        genero = "O"
    Else
    
        For Each group In oExUser.GetMemberOfList
            If (InStr(UCase(group.Name), "MASCULINO") > 0) Then
                 genero = "H"
                 Exit For
            ElseIf (InStr(UCase(group.Name), "FEMENINO") > 0) Then
                 genero = "F"
                 Exit For
            Else
                 genero = "O"
                 'Exit For
            End If
        Next
    End If
    getGeneroDominio = genero
End Function

Private Function getCabecera(destinatarios As Recipients)
    Dim cabecera As String
    Dim nombre As String
    Dim oExUser As Outlook.ExchangeUser
    Dim myNamespace As Outlook.NameSpace
    Dim genero As String
    cabecera = "Estimad"
    If getCountDestinariosTo(destinatarios) = 1 Then
        Set myNamespace = Application.GetNamespace("MAPI")
        Set oExUser = myNamespace.GetAddressEntryFromID(destinatarios(1).EntryID).GetExchangeUser
        genero = getGeneroDominio(oExUser)
        If oExUser Is Nothing Then
            nombre = ""
        Else
            nombre = Split(oExUser.FirstName, " ")(0)
        End If
        
        If genero = "H" Then
            cabecera = cabecera & "o " & nombre
        ElseIf genero = "F" Then
            cabecera = cabecera & "a " & nombre
        Else
            cabecera = cabecera & "a/o"
        End If
    Else
        cabecera = cabecera & "os"
    End If
    
    cabecera = cabecera & ","
    getCabecera = cabecera
End Function

Private Function getSaludo()
    Dim saludo As String
    If Time >= #12:00:00 AM# And Time < #12:00:00 PM# Then
        saludo = "Buenos Días,"
    ElseIf Time >= #12:00:00 PM# And Time < #6:00:00 PM# Then
        saludo = "Buenas Tardes,"
    ElseIf Time >= #6:00:00 PM# And Time < #12:00:00 AM# Then
        saludo = "Buenas Noches,"
    End If
    
    getSaludo = saludo
End Function

Private Sub oItem_ReplyAll(ByVal Response As Object, Cancel As Boolean)
    'Set olReply = oItem.Reply
    Dim olInspector As Inspector
    Dim saludo As String
    
    Dim nombre As String
    Dim inicio As String
    'oItem.HTMLBody = "Hello, Thank you. " & vbCrLf & Response.HTMLBody
    'If Not olItem Is Nothing And olItem.Class = olMail Then
    'MsgBox "Fin"
    Set olInspector = Response.GetInspector
    If olInspector.EditorType = olEditorWord Then
        Set objWDoc = olInspector.WordEditor
        Set objWSel = objWDoc.Application.Selection
        
        saludo = getSaludo
        
        'Response.HTMLBody = "Estimad@, " & vbCrLf & saludo & vbCrLf & Response.HTMLBody
        Response.Display
        inicio = getCabecera(Response.Recipients)
        SendKeys inicio, True
        SendKeys "{ENTER}", True
        SendKeys saludo, True
        SendKeys "{ENTER}", True
        SendKeys "{ENTER}", True
        SendKeys "{ENTER}", True
        SendKeys "{ENTER}", True
        SendKeys "Saludos ", True
        SendKeys "{UP}", True
        SendKeys "{UP}", True
        
    End If
        
End Sub

Private Function getCountDestinariosTo(destinatarios As Recipients) As Integer
    Dim cont As Integer
    cont = 0
    For Each dest In destinatarios
        If dest.Type = olTo Then
            cont = cont + 1
        End If
    Next
    
    getCountDestinariosTo = cont
End Function
