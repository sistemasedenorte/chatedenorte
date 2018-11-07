Attribute VB_Name = "ModProtocol"
Option Explicit

Public Function ParseUserEntersLeaves(Packet As String) As Boolean
    'Packet is structured like this:
    '"ENT" & Chr$(2) & Nickname
    '"LEA" & Chr$(2) & Nickname
    ParseUserEntersLeaves = True
    If InStr(1, Packet, Chr$(2)) > 0 Then
        Dim strInfo() As String, bolEntering As Boolean
        
        strInfo = Split(Packet, Chr$(2))
        
        'para evitar nombres en blanco en la lista
        If Trim(ObtenerNombreCompleto(strInfo(1))) = "" Then
           Exit Function
        End If
        
        bolEntering = strInfo(0) = "ENT"
        AddUserEntersLeaves frmChat.rtbChat, strInfo(1), bolEntering
        
        If bolEntering Then
            If YaEstaDentro(strInfo(1)) = True Then
               ParseUserEntersLeaves = False
            End If
            frmChat.lstUsers.AddItem ObtenerNombreCompleto(strInfo(1))
        Else
            RemoveListItem frmChat.lstUsers, ObtenerNombreCompleto(strInfo(1))
        End If
        
    End If
    
End Function

Private Function YaEstaDentro(idRed As String) As Boolean
   Dim nombreUsuario As String, i As Long
   
     YaEstaDentro = False
     nombreUsuario = ObtenerNombreCompleto(idRed)
     For i = 0 To frmChat.lstUsers.ListCount - 1
         If frmChat.lstUsers.List(i) = nombreUsuario Then
             YaEstaDentro = True
             Exit For
         End If
     Next
    
End Function


Public Function ParseChatMessage(Packet As String) As Boolean
    
    ParseChatMessage = False
    'Packet is structured like this:
    '"MSG" & Chr$(2) & Nickname & Chr$(2) & Message
    If InStr(1, Packet, Chr$(2)) > 0 Then
        Dim strInfo() As String
        
        strInfo = Split(Packet, Chr$(2))
        'strInfo(0) = MSG
        'strInfo(1) = Nickname
        'strInfo(2) = Message
        
       'Solo recibo informaciones de gente que puede interactuar conmigo una
       'via u otra y del servidor
       If strInfo(1) = "Server" Or strMyNickname = "Server" Then
           AddChatMessage frmChat.rtbChat, strInfo(1), strInfo(2)
           Erase strInfo
           Exit Function
       End If
       
       'Lista de Puntos de Conversacion:
       'PuntoA: Elias
       'PuntoB: Victoria
       'Comparar PuntoA con PuntoB y viceversa
'       If (strInfo(1) = "Victoria" And strMyNickname = "Elias") Or _
'        (strInfo(1) = "Elias" And strMyNickname = "Victoria") Or _
'         (strInfo(1) = "Victoria" And strMyNickname = "Victoria") Or _
'         (strInfo(1) = "Elias" And strMyNickname = "Elias") Then
'          AddChatMessage frmChat.rtbChat, strInfo(1), strInfo(2)
'       Else
'         ' MsgBox "No se pudo: " & strInfo(1) & " Nick " & strMyNickname, vbCritical, "Lo siento"
'       End If
       If gListaPermisos Is Nothing Then CargarListaPermisos
       If InStr(1, Trim(strInfo(2)), "@" & Trim(ObtenerNombreCompleto(strMyNickname))) = 0 And Trim(strInfo(1)) <> strMyNickname Then
           Erase strInfo
           Exit Function 'si el mensaje no es para mi
       End If
       
       ParseChatMessage = True
       If EstaEnLaListaLaPareja(Trim(strInfo(1)), strMyNickname, gListaPermisos) = True Then
          AddChatMessage frmChat.rtbChat, Trim(strInfo(1)), Trim(strInfo(2))
       End If
       'Fin lista de puntos de conversacion
       
       Erase strInfo
       
    End If
    
End Function

'Aqui voy a leer la lista de grupos y sus permisos
'no leere los usuarios en si, solo los grupos correspondientes
Public Function EstaEnLaListaLaPareja(ByVal personaA As String, ByVal personaB As String, lista As Collection) As Boolean
    Dim datosItem, item
    Dim GrupoA As String, grupoB As String
    Dim laMismaPersona As Boolean
    
    
    EstaEnLaListaLaPareja = False
    laMismaPersona = False
    If Trim(personaA) = Trim(personaB) Then laMismaPersona = True
    GrupoA = ObtenerGrupoRelacionado(Trim(personaA))
    grupoB = ObtenerGrupoRelacionado(Trim(personaB))
    
    'Si estan en el mismo grupo
    'no pueden hablar entre si a menos que el nombre del grupo
    'tengan asterisco
'    If GrupoA = grupoB And InStr(1, GrupoA, "*") = 0 Then
'       EstaEnLaListaLaPareja = False
'       Exit Function
'    End If
'
'
    'Como evaluare los grupos mas no personas
    'entonces hago el cambio
    personaA = GrupoA
    personaB = grupoB
    
    'aqui valido que este en la lista de parejas validas
    'para que el mensaje se le envien al chat del emisor
    If personaA = personaB Then
        For Each item In lista
            
            'Debug.Print item
            If InStr(1, item, ",") > 0 Then
                datosItem = Split(item, ",")
                datosItem(0) = Replace(datosItem(0), vbTab, "")
                datosItem(1) = Replace(datosItem(1), vbTab, "")
                If Trim(personaA) = Trim(CStr(datosItem(0))) Or Trim(personaA) = Trim(CStr(datosItem(1))) Then
                   'Entre si pueden hablar si el nombre del grupo tiene asterisco
                   EstaEnLaListaLaPareja = False
                   If InStr(1, GrupoA, "*") > 0 Then EstaEnLaListaLaPareja = True
                   If laMismaPersona = True And GrupoA = grupoB Then EstaEnLaListaLaPareja = True
                   Exit For
                End If
             End If
             
        Next
    Else
        For Each item In lista
           If InStr(1, item, ",") > 0 Then
             datosItem = Split(item, ",")
             datosItem(0) = Replace(datosItem(0), vbTab, "")
             datosItem(1) = Replace(datosItem(1), vbTab, "")
             If Trim(personaA) = Trim(CStr(datosItem(0))) And Trim(personaB) = Trim(CStr(datosItem(1))) Or _
                Trim(personaB) = Trim(CStr(datosItem(0))) And Trim(personaA) = Trim(CStr(datosItem(1))) Then
                 EstaEnLaListaLaPareja = True
                Exit For
             End If
           End If
        Next
    End If
    
    
End Function


Private Sub CargarListaPermisos()
   Dim fs As New FileSystemObject
   Dim archivo As TextStream
   Dim linea As String
   
   On Error GoTo fixme
   'servidor:127.0.0.1
   'puerto:1234

   If Not gListaPermisos Is Nothing Then Exit Sub
   Set gListaPermisos = New Collection
   Set archivo = fs.OpenTextFile("\\" & gServidor & "\LogChatEdenorte\Permisos.txt", ForReading)
   Do While archivo.AtEndOfStream = False
      linea = archivo.ReadLine
      gListaPermisos.Add linea
   Loop
   archivo.Close
   Set fs = Nothing
   
   Exit Sub
   
fixme:
   MsgBox "No existe el archivo de permisos", vbCritical, "Falta un archivo"
   
End Sub

'Me recupera las informaciones de una persona logeada al chat
Public Function ObtenerNombreCompleto(idDeRed As String) As String
   Dim fs As New FileSystemObject
   Dim archivo As TextStream
   Dim linea, datos
   
   CargarListaPermisos
   ObtenerNombreCompleto = ""
   For Each linea In gListaPermisos
      If InStr(1, linea, idDeRed) > 0 Then
        datos = Split(linea, vbTab)
        ObtenerNombreCompleto = CStr(datos(0)) & " " & CStr(datos(1)) '& " " & CStr(datos(2))
        Exit For
      End If
   Next
   ObtenerNombreCompleto = Trim(ObtenerNombreCompleto)
   
End Function


Public Function ObtenerIdRed(Codigo As String) As String
   Dim fs As New FileSystemObject
   Dim archivo As TextStream
   Dim linea, datos
   
   CargarListaPermisos
   ObtenerIdRed = ""
   For Each linea In gListaPermisos
      If InStr(1, linea, Codigo) > 0 Then
        datos = Split(linea, vbTab)
        ObtenerIdRed = CStr(datos(3)) 'hoy 3
        Exit For 'hoy
      End If
   Next
   
End Function


'Me dice a cual  grupo pertenece una persona, id de red
Public Function ObtenerGrupoRelacionado(idDeRed As String) As String
   Dim fs As New FileSystemObject
   Dim archivo As TextStream
   Dim linea, grupo As String
   
   CargarListaPermisos
   ObtenerGrupoRelacionado = ""
   For Each linea In gListaPermisos
      
      If InStr(1, linea, "<") > 0 And InStr(1, linea, "/") = 0 Then
         grupo = Replace(linea, "<", "")
         grupo = Replace(grupo, ">", "")
      End If
      
      If InStr(1, linea, idDeRed) > 0 Then
         ObtenerGrupoRelacionado = grupo
         Exit For
      End If
   Next
   
End Function


Public Sub ParseUserList(Packet As String)
    'Packet is structured like this:
    '"LST" & Chr$(2) & User1 & vbCrLf & User2 & vbCrLf & User3
    If InStr(1, Packet, Chr$(2)) > 0 Then
        Dim strInfo() As String, strUsers() As String
        Dim intLoop As Integer
        
        strInfo() = Split(Packet, Chr$(2))
        'strInfo(0) = LST
        'strInfo(1) = User1 & vbCrLf & User2 & vbCrLf & User3
        With frmChat.lstUsers
            .Clear
            
            If InStr(1, strInfo(1), vbCrLf) > 0 Then
                strUsers = Split(strInfo(1), vbCrLf)
                
                For intLoop = 0 To UBound(strUsers)
                    If Len(strUsers(intLoop)) > 0 Then
                        .AddItem ObtenerNombreCompleto(strUsers(intLoop))
                    End If
                Next intLoop
            Else
                .AddItem ObtenerNombreCompleto(strInfo(1))
            End If
            
        End With
        
        Erase strInfo
        Erase strUsers
    End If
    
End Sub
