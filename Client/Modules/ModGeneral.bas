Attribute VB_Name = "ModGeneral"
'modGeneral.bas
'--------------
'General subs/functions used throughout the program.
'These are simple things like adding a message to
'a RichTextBox, and other things that aren't really
'relevant/necessary for this chat example.
Option Explicit

'Program is ending?
Public bolEnding As Boolean


Public Sub EndProgram()
    If Not bolEnding Then
        bolEnding = True
        frmChat.sckClient.Close
        Unload frmChat
        Unload frmConnect
    End If
End Sub

'Removes a ListItem from a ListBox.
Public Sub RemoveListItem(ListObject As Object, _
    ItemText As String, Optional ByVal RemoveAll As Boolean = False)
    
    'Make sure ListObject is a ListBox or ComboBox.
    If Not TypeOf ListObject Is ListBox And Not _
           TypeOf ListObject Is ComboBox Then Exit Sub
           
    Dim intLoop As Integer, strL As String
    
    'Compare lowercase values (case-insensitive search).
    strL = LCase$(ItemText)
    
    With ListObject
        'Loop through all ListItems.
        For intLoop = 0 To .ListCount - 1
            'If LowerCase(this item) equals item we're searching for...
            If LCase$(.List(intLoop)) = strL Then
                .RemoveItem intLoop
                If Not RemoveAll Then Exit For
            End If
        Next intLoop
    End With
    
End Sub

'Checks if an item exists in a ListBox.
'If it does, it returns the ListIndex.
Public Function FindListItem(ListObject As Object, _
    ItemText As String) As Integer
    
    'Make sure ListObject is a ListBox or ComboBox.
    If Not TypeOf ListObject Is ListBox And Not _
           TypeOf ListObject Is ComboBox Then Exit Function
           
    Dim intLoop As Integer, strL As String
    
    'Compare lowercase values (case-insensitive search).
    strL = LCase$(ItemText)
    
    With ListObject
        'Loop through all ListItems.
        For intLoop = 0 To .ListCount - 1
            'If LowerCase(this item) equals item we're searching for...
            If LCase$(.List(intLoop)) = strL Then
                'Return the index of the item.
                FindListItem = intLoop
                'Exit.
                Exit For
            End If
        Next intLoop
    End With
    
End Function

'Adds a status message to the chat RTB.
Public Sub AddStatusMessage(RTB As Object, _
    ByVal Color As Long, Message As String)
    
    'Make sure RTB is a RichTextBox.
    If Not TypeOf RTB Is RichTextBox Then Exit Sub
    
    With RTB
        .SelStart = Len(.Text) 'Move to end of RTB.
        .SelFontName = "Tahoma" 'Set font.
        .SelFontSize = 11 'Set font size.
        .SelBold = False
        .SelItalic = False
        .SelUnderline = False
        .SelColor = Color
        .SelText = Message & vbCrLf
    End With
    
End Sub

'Displays a user entering/leaving message.
Public Sub AddUserEntersLeaves(RTB As Object, _
    userName As String, Optional ByVal Entering As Boolean = True)
    
    If Not TypeOf RTB Is RichTextBox Then Exit Sub
    
    Dim strL As String
    
    strL = LCase$(userName)
    
    With RTB
        .SelStart = Len(.Text)
        .SelFontName = "Tahoma"
        .SelFontSize = 11
        .SelBold = True
        .SelItalic = False
        .SelUnderline = False
        .SelColor = IIf(strL = LCase$(strMyNickname), RGB(0, 0, 128), RGB(128, 0, 0))
        .SelText = userName & " "
        .SelBold = False
        .SelColor = IIf(Entering, RGB(0, 128, 0), RGB(128, 128, 128))
        .SelText = IIf(Entering, "has joined the room.", "has left the room.") & vbNewLine
    End With
    
End Sub

'Displays a chat message.
Public Sub AddChatMessage(RTB As Object, _
    userName As String, Message As String)
    
    If Not TypeOf RTB Is RichTextBox Then Exit Sub
    
    Dim strL As String
    
    strL = LCase$(userName)
    
    'solo puedo ver mensajes enviados a mi: 8-dic-2015
'    If InStr(1, Message, "@" & ObtenerNombreCompleto(strMyNickname)) = 0 _
'        And Username <> strMyNickname Then
'        Exit Sub
'    End If
    
    With RTB
        .SelStart = Len(.Text)
        .SelFontName = "Tahoma"
        .SelFontSize = 11
        .SelBold = True
        .SelItalic = False
        .SelUnderline = False
        .SelColor = IIf(strL = LCase$(strMyNickname), RGB(0, 0, 128), RGB(128, 0, 0))
        .SelText = userName & ": "
        .SelBold = False
        .SelColor = vbBlack
        .SelText = Message & vbCrLf
    End With
    
    ClasificarMensaje Message, userName
    
End Sub

Private Sub ClasificarMensaje(mensaje As String, userName As String)
   Dim item As cTab, contenido As RichTextBox, ExisteElTab As Boolean
   Dim i As Long, strL As String, nombreEmisor As String, lineaFinal As String
   

   strL = LCase$(userName)
   nombreEmisor = "": ExisteElTab = False
   
   If strL <> LCase$(strMyNickname) Then 'Mensaje entrante
      nombreEmisor = Trim(ObtenerNombreCompleto(userName))
      
      i = 0
      Do While i <= frmChat.tabPrincipal.Tabs.Count - 1
           i = i + 1
           Set item = frmChat.tabPrincipal.Tabs.item(i)
           
            If InStr(1, nombreEmisor, item.Caption) > 0 Then
               ExisteElTab = True
               Exit Do
            End If
      Loop
      
      'Agrego un nuevo tab si no existe:
      If ExisteElTab = False Then
        frmChat.AgregarPestana Trim(nombreEmisor)
      End If
      
   End If
   
   'Mensajes enviados
   i = 0
   Do While i <= frmChat.tabPrincipal.Tabs.Count - 1
      i = i + 1
      Set item = frmChat.tabPrincipal.Tabs.item(i)
      
      'Las respuestas vienen dadas porque el id de red es diferente al actual
      'tengo abierto el tab de quien le envio 111-JUANA
      'ese tab debe de alimentarse si me envian respuesta a mi
      
      If InStr(1, mensaje, item.Caption) > 0 Or _
         InStr(1, nombreEmisor, item.Caption) > 0 Then
           
           Set contenido = item.Panel
           lineaFinal = Replace(mensaje, "@" & item.Caption, "")
           lineaFinal = Replace(lineaFinal, "@" & nombreEmisor, "")
           lineaFinal = Replace(lineaFinal, "@" & Trim(ObtenerNombreCompleto(strMyNickname)), "")
           With contenido
                .SelStart = Len(.Text)
                .SelFontName = "Tahoma"
                .SelFontSize = 11
                .SelBold = True
                .SelItalic = False
                .SelUnderline = False
                .SelColor = IIf(strL = LCase$(strMyNickname), RGB(0, 0, 128), RGB(128, 0, 0))
                .SelText = userName & ": "
                .SelBold = False
                .SelColor = vbBlack
                .SelText = lineaFinal & vbCrLf
           End With
           
           'aqui si el usuario es diferente a logueado, es una respuesta
           'por tanto notifico del mensaje con un icono amarillo
           If strL <> LCase$(strMyNickname) Then
               item.IconIndex = 5
           End If
           
      End If
       
   Loop
   
End Sub

