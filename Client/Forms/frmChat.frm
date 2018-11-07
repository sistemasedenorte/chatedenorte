VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbaliml6.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "ARGradient.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{546C0534-0DE2-457D-ACB3-531B0833BC86}#1.0#0"; "VB Splitter.ocx"
Begin VB.Form frmChat 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Cliente EdenorteCHAT"
   ClientHeight    =   7380
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   492
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   638
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   3300
      Left            =   5400
      TabIndex        =   6
      Top             =   1800
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   5821
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmChat.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2400
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Indicar Archivo Chat"
      Filter          =   "Chats|*.rtf"
      InitDir         =   "C:\LogChatEdenorte"
   End
   Begin VB.PictureBox picTitulo 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin ARGradientControl.ARGradient ARGradient1 
         Height          =   735
         Left            =   1080
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1296
         Color           =   32768
         FinColor        =   14737632
         Orientation     =   4
         Caption         =   "Buenas"
         ShowCaption     =   -1  'True
         ForeColor       =   16777215
         Alignment       =   0
         VerticalAlignment=   0
         GradientSteps   =   100
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "frmChat.frx":0946
         Top             =   0
         Width           =   480
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   1270
      ButtonWidth     =   1455
      ButtonHeight    =   1217
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refrescar"
            Key             =   "refrescar"
            Object.ToolTipText     =   "Refrescar - F5"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Soporte"
            Key             =   "soporte"
            Object.ToolTipText     =   "Contactar Soporte F1"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mensaje"
            Key             =   "mensajegrupo"
            Object.ToolTipText     =   "Enviar Mensaje a Grupo  CTRL+G"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar Conversacion Actual CTRL+S"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del Sistema CTRL+X"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Index           =   0
         Left            =   5160
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmChat.frx":0EDB
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmChat.frx":15D5
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmChat.frx":1CCF
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmChat.frx":23C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmChat.frx":2AC3
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmChat.frx":31BD
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VBSplitter.Splitter Splitter1 
      Height          =   5820
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   10266
      FillContainer   =   0   'False
      Begin vbalDTab6.vbalDTabControl tabPrincipal 
         Height          =   4596
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   6025
         _ExtentX        =   10636
         _ExtentY        =   8096
         TabAlign        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ListBox lstUsers 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4380
         IntegralHeight  =   0   'False
         Left            =   6085
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   0
         Width           =   3485
      End
      Begin VB.TextBox txtMsg 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1164
         Left            =   0
         MaxLength       =   1024
         TabIndex        =   3
         Top             =   4656
         Width           =   9570
      End
   End
   Begin ChamaleonButton.ChameleonBtn cmdSend 
      Height          =   735
      Left            =   7440
      TabIndex        =   1
      Top             =   6600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Enviar"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   0
      MPTR            =   1
      MICON           =   "frmChat.frx":370D
      PICN            =   "frmChat.frx":3729
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   5640
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin vbalIml6.vbalImageList ilsIcons 
      Left            =   600
      Top             =   6720
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   8036
      Images          =   "frmChat.frx":407D
      Version         =   131072
      KeyCount        =   7
      Keys            =   "ÿÿÿÿÿÿ"
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuArchivoEnviar 
         Caption         =   "Enviar Mensaje"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuArchivoMensajeGrupo 
         Caption         =   "Enviar Mensaje a Grupo..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuArchivoSalvar 
         Caption         =   "Salvar Conversación..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuArchivoCargarConv 
         Caption         =   "Cargar Conversacion..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuArchivoRefrescar 
         Caption         =   "Refrescar"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuArchivoLinea 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mnuAyudaSoporte 
         Caption         =   "Contactar Soporte"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAyudaAcercaDe 
         Caption         =   "Acerca De"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_layout As LayoutManager.DynamicLayout
Private WithEvents m_frmSysTray As frmSysTray
Attribute m_frmSysTray.VB_VarHelpID = -1

'JUST USED FOR FORM RESIZING.
Private Type RECT
    rctLeft As Long
    rctTop As Long
    rctRight As Long
    rctBottom As Long
End Type

'JUST USED FOR FORM RESIZING.
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'JUST USED FOR FORM RESIZING.
Private udtMyRect As RECT

'Received data buffer.
Private strBuffer As String

Private Sub cmdSend_Click()
   
    If Len(txtMsg.Text) > 0 Then
        If sckClient.State <> sckConnected Then
            AddStatusMessage rtbChat, RGB(128, 0, 0), "> Not connected! Cannot send message."
        Else
            Dim strPacket As String
            
            If TienenPermisos = False Then
               AddStatusMessage rtbChat, RGB(128, 0, 0), "> No Puede conversar con esta persona"
               
               Exit Sub
            End If
            
            strPacket = "MSG" & Chr$(2) & strMyNickname & Chr$(2) & lstUsers.List(lstUsers.ListIndex) & Chr$(2) & txtMsg.Text & "@" & lstUsers.List(lstUsers.ListIndex) & Chr$(4)
            sckClient.SendData strPacket
            txtMsg.Text = ""
        End If
    End If
    
End Sub


Private Sub EnviarMensajeGrupo(nombreGrupo)
   Dim i As Long, usuario As String
   Dim leEnvioaAlguien As Boolean
   Dim strPacket As String
   Dim datos, idRed As String
   
   
   leEnvioaAlguien = False
   For i = 0 To (lstUsers.ListCount - 1)
        lstUsers.ListIndex = i
        
        'le enviara el mensaje a todos los de  grupo seleccionado:
        '1-Si tiene los permisos para hablar con ese grupo
        '2-Si cada destinatario pertenece a ese grupo
        '3-Si el destinatario no es el usuario actual
        datos = Split(lstUsers.List(lstUsers.ListIndex), " ")
        idRed = ObtenerIdRed(CStr(datos(0)))
        If TienenPermisos = True And lstUsers.List(lstUsers.ListIndex) _
            <> ObtenerNombreCompleto(strMyNickname) _
            And ObtenerGrupoRelacionado(idRed) = nombreGrupo Then
            
            strPacket = "MSG" & Chr$(2) & strMyNickname & Chr$(2) & lstUsers.List(lstUsers.ListIndex) & Chr$(2) & txtMsg.Text & "@" & lstUsers.List(lstUsers.ListIndex) & Chr$(4)
            sckClient.SendData strPacket
            leEnvioaAlguien = True
            
        End If
   Next
   
   If leEnvioaAlguien = False Then
      AddStatusMessage rtbChat, RGB(128, 0, 0), "> No se envió el mensaje al grupo: No están logueados en el Chat o debe revisar sus privilegios"
   Else
      AddStatusMessage rtbChat, RGB(181, 230, 29), "> Mensaje enviado exitosamente al grupo " & nombreGrupo
   End If

End Sub


Private Function TienenPermisos() As Boolean
   Dim datos, idRed As String
   
   datos = Split(lstUsers.List(lstUsers.ListIndex), " ")
   idRed = ObtenerIdRed(CStr(datos(0)))
   
   TienenPermisos = EstaEnLaListaLaPareja(strMyNickname, idRed, gListaPermisos)
   
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If sckClient.State <> sckConnected Then
            frmConnect.Show
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
   Dim fs As New FileSystemObject
   Dim item As cTab
   
   frmChat.Caption = "Bienvenid@s!!"
   tabPrincipal.ImageList = ilsIcons
     
  ' lblTitulo.Caption = "[" & ObtenerNombreCompleto(strMyNickname) & "]"
   ARGradient1.Caption = "[" & ObtenerNombreCompleto(strMyNickname) & "]"
   ARGradient1.FinColor = Me.BackColor
   
   Me.Refresh
   Splitter1.Visible = True
   Set m_layout = New LayoutManager.DynamicLayout
   m_layout.Insert Splitter1, apAll
   m_layout.Insert cmdSend, apBottom Or apRight
   m_layout.Insert picTitulo, apRight
   
  Set m_frmSysTray = New frmSysTray
    With m_frmSysTray
        .AddMenuItem "&Open SysTray Sample", "open", True
        .AddMenuItem "-"
        .AddMenuItem "&Close", "close"
        .ToolTip = "SysTray Sample!"
        .IconHandle = Me.Icon.Handle
    End With
    
    If fs.FolderExists("C:\LOGCHATEDENORTE") = False Then
       fs.CreateFolder ("C:\LOGCHATEDENORTE")
    End If
    
    Set item = tabPrincipal.Tabs.Add(, , "General")
    Set item.Panel = rtbChat
    
End Sub


Private Sub Reconectar()

   With sckClient
        .Close
        bolRecon = False
        .Connect
    End With
    AddStatusMessage frmChat.rtbChat, RGB(128, 128, 128), "> Connecting to " & sckClient.RemoteHost & ":" & sckClient.RemotePort & "..."

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload m_frmSysTray
    Set m_frmSysTray = Nothing
End Sub


Private Sub lstUsers_DblClick()
  
  AgregarPestana lstUsers.List(lstUsers.ListIndex)
   
End Sub


Public Sub AgregarPestana(titulo As String)
    Dim item As RichTextBox, item2 As cTab, i As Long, ExisteElTab As Boolean
      
    i = 0
    ExisteElTab = False
    'Si existe solo
    Do While i <= frmChat.tabPrincipal.Tabs.Count - 1
        i = i + 1
        Set item2 = frmChat.tabPrincipal.Tabs.item(i)
        
        If InStr(1, titulo, item2.Caption) > 0 Then
            ExisteElTab = True
            item2.Selected = True
            Exit Sub
        End If
    Loop
    
    'Agrego un nuevo tab si no existe:
    Set item2 = tabPrincipal.Tabs.Add("chatbox" & tabPrincipal.Tabs.Count, , titulo, 0)
    
    item2.Selected = True
    item2.IconIndex = 2
    item2.CanClose = True
    
    'Agrego una caja de texto tambien
    Set item = Controls.Add("RICHTEXT.RichtextCtrl.1", "chatbox" & tabPrincipal.Tabs.Count - 1, Me)
    Set item2.Panel = item
    item.Text = "": item.AutoVerbMenu = True
  
End Sub


Private Sub m_frmSysTray_BalloonClicked()
   Dim Form, i As Long
   
   i = 0
   For Each Form In Forms
       Forms(i).ZOrder 1
       i = i + 1
   Next
   
End Sub

Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
    Select Case sKey
    Case "open"
        Me.Show
        Me.ZOrder
    Case "close"
        Unload Me
    End Select
    
End Sub

Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    Me.Show
    Me.ZOrder
End Sub

Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    If (eButton = vbRightButton) Then
        m_frmSysTray.ShowMenu
    End If
End Sub


'Form is resizing.
'-----------------
'Resizes according to the CLIENT AREA of the form.
'Form.Width/Form.Height/Form.ScaleWidth/Form.ScaleHeight return width + non-client area (borders, etc.).
'This provides pixel perfect resizing regardless of which windows theme/screen resolution is being used.
Private Sub Form_Resize()
    'Don't do anything if form is being minimized.
'    If Me.WindowState = vbMinimized Then Exit Sub
    
    GetClientRect Me.hwnd, udtMyRect
    m_layout.Resize

'    rtbChat.Width = udtMyRect.rctRight - lstUsers.Width
'    rtbChat.Height = udtMyRect.rctBottom - picTitulo.Height - 50
'    lstUsers.Height = rtbChat.Height - picTitulo.Height
'    lstUsers.Left = rtbChat.Width + 15
'    cmdSend.Top = rtbChat.Height + 50
'    cmdSend.Left = lstUsers.Left
'    cmdSend.Width = lstUsers.Width
'
'    Splitter1.Width = Me.Width
'    Splitter1.Height = Me.Height - cmdSend.Height
    
'    txtMsg.Top = cmdSend.Top
'    txtMsg.Width = rtbChat.Width
'    lblTitulo.Width = Me.Width
'    picTitulo.Width = Me.Width
    
End Sub


'End program.
Private Sub Form_Unload(Cancel As Integer)
   If Not bolRecon Then
        EndProgram
    End If
End Sub

Private Sub lstUsers_Click()

  cmdSend.Enabled = False
  If lstUsers.ListIndex > -1 Then
  
     If lstUsers.List(lstUsers.ListIndex) <> ObtenerNombreCompleto(strMyNickname) Then
        cmdSend.Enabled = True
     End If
     
   End If

End Sub


Private Sub mnuArchivoCargarConv_Click()
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen

    If CommonDialog1.FileName <> "" Then
      rtbChat.LoadFile CommonDialog1.FileName
    End If
    
End Sub

Private Sub mnuArchivoEnviar_Click()
  If cmdSend.Enabled = True Then cmdSend_Click
End Sub

Private Sub mnuArchivoMensajeGrupo_Click()
    
    If txtMsg.Text = "" Then
       MsgBox "Debe de indicar el mensaje a enviar", vbExclamation, "Enviar Mensaje a Grupo"
       Exit Sub
    End If
    
    frmSeleccionarGrupo.Show vbModal
    If frmSeleccionarGrupo.g_OK = False Then Exit Sub
    
    EnviarMensajeGrupo frmSeleccionarGrupo.g_nombreGrupo
    
End Sub

Private Sub mnuArchivoRefrescar_Click()
    Reconectar
   ' MsgBox "Datos refrescados", vbInformation, "Refrescar"
End Sub

Private Sub mnuArchivoSalvar_Click()
  Dim nombre As String
  
  nombre = "C:\LogChatEdenorte\Chat" & UCase(strMyNickname) & "-" & Format(Now, "DDMMMYYYY_hh") & ".rtf"
  rtbChat.SaveFile nombre
  MsgBox "Archivo Salvado en " & nombre, vbInformation, "Salvar conversación"
  
End Sub

Private Sub mnuAyudaSoporte_Click()
   MsgBox "Favor llamar al 5004 o envíe un correo a cgs@edenorte.com.do", vbInformation, "Soporte"
End Sub

Private Sub mnuSalir_Click()
  EndProgram
End Sub

'Contents of rtbChat have changed.
Private Sub rtbChat_Change()
    'Auto-scroll box.
    rtbChat.SelStart = Len(rtbChat.Text)
    'Sets cursor position (carrot) to end (length of text) of RTB.
End Sub


'The server has closed the connection! Connection lost.
'------------------------------------------------------
Private Sub sckClient_Close()
    sckClient.Close
    bolRecon = True
    AddStatusMessage rtbChat, RGB(128, 128, 128), "> The connection to the server was lost! Press the [ESC] key to re-connect."
End Sub

Private Sub sckClient_Connect()
    AddStatusMessage rtbChat, RGB(0, 128, 0), "> Connected!"
    
    'Send nickname to server.
    Dim strPacket As String
    
    strPacket = "CON" & Chr$(2) & strMyNickname & Chr$(4)
    sckClient.SendData strPacket
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String, strPackets() As String
    Dim strTrunc As String, bolTrunc As Boolean
    Dim lonLoop As Long, lonTruncStart As Long
    Dim lonUB As Long
    
    sckClient.GetData strData, vbString, bytesTotal
    strBuffer = strBuffer & strData
    strData = vbNullString
    
    If Right$(strBuffer, 1) <> Chr$(4) Then
        bolTrunc = True
        lonTruncStart = InStrRev(strBuffer, Chr$(4))
        If lonTruncStart > 0 Then
            strTrunc = Mid$(strBuffer, lonTruncStart + 1)
        End If
    End If
    
    If InStr(1, strBuffer, Chr$(4)) > 0 Then
        strPackets() = Split(strBuffer, Chr$(4))
        lonUB = UBound(strPackets)
        
        If bolTrunc Then lonUB = lonUB - 1
        
        For lonLoop = 0 To lonUB
            If Len(strPackets(lonLoop)) > 3 Then
                
                Select Case Left$(strPackets(lonLoop), 3)
                    
                    'Packet is a chat message.
                    Case "MSG"
                        If ParseChatMessage(strPackets(lonLoop)) = True Then
                           Beep
                           Beep
                           m_frmSysTray.ShowBalloonTip "Hola! tienes un nuevo mensaje :)", "Nuevo Mensaje", NIIF_INFO
                        End If
 
                    'User list has been sent.
                    Case "LST"
                        ParseUserList strPackets(lonLoop)
                    
                    Case "ENT", "LEA"
                       
'                       If ParseUserEntersLeaves(strPackets(lonLoop)) = False Then
'                            MsgBox "Ya su usuario está dentro con otra sesión/pc" & _
'                            vbCrLf & "Favor usar una sola sesión", vbExclamation, "Usuario Duplicado"
'                            'Unload Me
'                       End If
                        
                    'Add your own here! :)
                    'Case "XXX"
                        'Do something.
                    
                    'Case "YYY"
                        'Do something.
                        
                End Select
            End If
        Next lonLoop
    
    End If
    
    Erase strPackets
    
    strBuffer = vbNullString
    
    If bolTrunc Then
        strBuffer = strTrunc
    End If
    
    strTrunc = vbNullString
End Sub


Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckClient.Close
    bolRecon = True
    AddStatusMessage rtbChat, RGB(128, 0, 0), "> Error (" & Number & "): " & Description & IIf(Right$(Description, 1) = ".", "", ".")
    AddStatusMessage rtbChat, RGB(128, 0, 0), "> Press the [ESC] key to re-connect."
End Sub


Private Sub tabPrincipal_TabClose(theTab As vbalDTab6.cTab, bCancel As Boolean)
    Controls.Remove Controls(theTab.Key)
End Sub

Private Sub tabPrincipal_TabSelected(theTab As vbalDTab6.cTab)
   Dim i As Long
   
   If theTab.Caption <> "General" Then theTab.IconIndex = 2
   
   'ahora seleciono de la lista el item correcto:
   For i = 0 To lstUsers.ListCount - 1
      If Trim(lstUsers.List(i)) = Trim(theTab.Caption) Then
         lstUsers.ListIndex = i
         Exit For
      End If
   Next
   
   
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
       Case "refrescar"
             mnuArchivoRefrescar_Click
       Case "salir"
            mnuSalir_Click
       Case "soporte"
            mnuAyudaSoporte_Click
       Case "mensajegrupo"
            mnuArchivoMensajeGrupo_Click
       Case "salvar"
            mnuArchivoSalvar_Click
   End Select
End Sub


Private Sub txtMsg_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cmdSend.Enabled = False Then Exit Sub
        cmdSend_Click
        KeyAscii = 0 'Gets rid of 'beep' sound.
    End If
End Sub
