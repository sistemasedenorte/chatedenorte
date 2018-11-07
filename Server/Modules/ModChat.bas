Attribute VB_Name = "ModChat"
Option Explicit

'Structure to store info about a user.
'Add your own stuff if you want.
Public Type CHAT_USER
    strNickname As String 'Nickname.
    strIP As String 'IP address.
    strBuffer As String 'Received data buffer from this client.
    
    'Add your own stuff like:
    'strRoomName As String 'For multiple rooms.
    'strTimeConnected As String 'To store what time they connected?
    'All code in the server is based off the info stored here.
End Type

'An array that contains all info about every user.
'A user's index in this array corresponds to the Winsock control responsible for this connection.
Public udtUsers() As CHAT_USER

'Max integer value, therefore, max simultaneous connections.
'(Most computers can't handle anywhere near this many).
Private Const MAX_INT As Integer = 32767

'One main sub that closes the server.
'Should be called before opening the server.
Public Sub CloseServer()
    'Steps:
    '------
    '1. Unload & close all Winsock controls.
    '2. Erase udtUsers() array to clear up memory.
    
    Dim intLoop As Integer
    
    With frmChat
        .sckServer(0).Close 'Close first control.
        
        If .sckServer.UBound > 0 Then
            'More than one Winsock control in the array.
            'Loop through and close/unload all of them.
            For intLoop = 1 To .sckServer.UBound
                .sckServer(intLoop).Close
                Unload .sckServer(intLoop)
            Next intLoop
        End If
    
    End With
    
    'Erase all current users from memory.
    Erase udtUsers
End Sub

'Finds an available Winsock control to use for an incoming connection.
'You can just copy/paste this code into your chat program if you want.
'Just change "sckServer" to the name of your Winsock control (array).
'And change MAX_INT to max simultaneous connections that you want (it is at top of this module).
Public Function NextOpenSocket() As Integer
    Dim intLoop As Integer, intFound As Integer
    
    With frmChat
        'First, see if there is only one Winsock control.
        If .sckServer.UBound = 0 Then
            'Just load #1.
            Load .sckServer(1)
            .sckServer(1).Close
            NextOpenSocket = 1
        Else
            'There is more than 1.
            'Loop through all of them to find one not being used.
            'If it is not being used, it's state will = sckClosed (no connections).
            For intLoop = 1 To .sckServer.UBound
                If .sckServer(intLoop).State = sckClosed Then
                    'Found one not being used.
                    intFound = intLoop
                    Exit For
                End If
            Next intLoop
            
            'Check if we found one.
            If intFound > 0 Then
                NextOpenSocket = intFound
            Else
                'Didn't find one.
                'Load a new one.
                'Unless we reached MAX_INT
                'which is max number of clients.
                If .sckServer.UBound + 1 < MAX_INT Then
                    'There is room for another one.
                    intFound = .sckServer.UBound + 1
                    Load .sckServer(intFound)
                    .sckServer(intFound).Close
                    NextOpenSocket = intFound
                Else
                    'Server is full!
                    Debug.Print "CONNECTION REJECTED! MAX CLIENTS (" & MAX_INT & ") REACHED!"
                End If
            
            End If
        End If
    End With
    
End Function

'Returns the upper bounds (UBound) of udtUsers array without an error.
Public Function UBUsers() As Long
    On Error GoTo ErrorHandler
    
    UBUsers = UBound(udtUsers)
    
    Exit Function
    
ErrorHandler:
    
End Function

'Send data
'Sends data to every connected client.
Public Sub SendGlobalData(Data As String)
    Dim intLoop As Integer
    Dim datosRecibidos, datosItem

    On Error GoTo ErrorHandler
    
    If gListaPermisos Is Nothing Then CargarListaPermisos
    
    'La data solo sera recibida por pares de clientes definidos :) entre Elias y Victoria
    'recorrer la lista de pares hasta encontrar que lo que recibimos
    'esta definido previamente
    datosRecibidos = Split(Data, Chr(2))
        
    If UBound(datosRecibidos) > 1 And InStr(1, Data, "Server") = 0 Then
        If CStr(datosRecibidos(0)) = "MSG" Then
           If EstaEnLaListaLaPareja(CStr(datosRecibidos(1)), CStr(datosRecibidos(2)), gListaPermisos) = True Then
               'CStr(datosRecibidos(2)) <> "Elias" And CStr(datosRecibidos(2)) <> "Victoria" Then Exit Sub
               'No esta definido en la lista de pres
               
    '           If CStr(datosRecibidos(2)) = "Elias" Then Data = Replace(Data, Chr(2) & "Elias", "")
    '           If CStr(datosRecibidos(2)) = "Victoria" Then Data = Replace(Data, Chr(2) & "Victoria", "")
               
               'Por alguna razon aqui estoy quitando el remitente????
               Data = Replace(Data, Chr(2) & Trim(CStr(datosRecibidos(2))), "")
               'Data = Data & "@" & Trim(CStr(datosRecibidos(2)))
           Else
               Data = datosRecibidos(0) & Chr(2) & datosRecibidos(1) & Chr(2) & datosRecibidos(2) & " **SIN PRIVILEGIOS**@" & datosRecibidos(1) & Chr(4)
           End If
        End If
    End If
    'Fin cambio... debo leer desde un file la relacion de conversaciones posibles
    
    With frmChat
        If .sckServer.UBound > 0 Then
            For intLoop = 1 To .sckServer.UBound
                .sckServer(intLoop).SendData Data
                DoEvents
            Next intLoop
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    'if err.Number = 40006 then 'Socket not connected.
    Resume Next
End Sub


Private Function EstaEnLaListaLaPareja(personaA As String, personaB As String, lista As Collection) As Boolean
    Dim datosItem, item
    
    EstaEnLaListaLaPareja = False
    For Each item In lista
         datosItem = Split(item, ",")
         If Trim(personaA) = Trim(CStr(datosItem(0))) And Trim(personaB) = Trim(CStr(datosItem(1))) Or _
            Trim(personaB) = Trim(CStr(datosItem(0))) And Trim(personaA) = Trim(CStr(datosItem(1))) Then
            EstaEnLaListaLaPareja = True
            Exit For
         End If
    Next
    
End Function


Private Sub CargarListaPermisos()
   Dim fs As New FileSystemObject
   Dim archivo As TextStream
   Dim linea As String
   
   'servidor:127.0.0.1
   'puerto:1234
   'permisos:LogChatEdenorte\Permisos.txt
   
   If fs.FolderExists("C:\LogChatEdenorte\") = False Then fs.CreateFolder ("C:\LogChatEdenorte\")
   If fs.FileExists("C:\LogChatEdenorte\Permisos.txt") = False Then fs.CreateTextFile ("C:\LogChatEdenorte\Permisos.txt")
   
   Set gListaPermisos = New Collection
   Set archivo = fs.OpenTextFile("C:\LogChatEdenorte\Permisos.txt", ForReading)
   Do While archivo.AtEndOfStream = False
      linea = archivo.ReadLine
      gListaPermisos.Add linea
   Loop
   archivo.Close
   Set fs = Nothing
   
End Sub
