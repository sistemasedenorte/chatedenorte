VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cliente EdenorteCHAT"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Conectar »"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtNickname 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Coloca Tu NickName"
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox txtPort 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3120
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "1234"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtServer 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   960
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Caption         =   "No Conectado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   3840
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nickname:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Port:"
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   3840
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim userName As String
Dim UserDomain As String


Private Sub cmdConnect_Click()
    'Check input.
    txtServer.Text = Trim$(txtServer.Text)
    txtPort.Text = Trim$(txtPort.Text)
    txtNickname.Text = Trim$(txtNickname.Text)
    
    If Len(txtServer.Text) = 0 Or Len(txtPort.Text) = 0 Or _
       Len(txtNickname.Text) = 0 Then
        
        MsgBox "Complete todos los campos!", vbCritical
        Exit Sub
    ElseIf Not IsNumeric(txtPort.Text) Then
        MsgBox "Numero de Puerto no valido!", vbCritical
        Exit Sub
    End If
    
    'Done with that...
    strMyNickname = txtNickname.Text
    If ObtenerNombreCompleto(strMyNickname) = "" Then
       MsgBox "Lo sentimos mucho pero al parecer usted " & vbCrLf & _
              " no tiene permisos definidos para usar el CHAT" & vbCrLf & _
              "------" & vbCrLf & _
              "Converse con el Administrador de esta" & vbCrLf & _
              "herramienta y digale que lo agregue como usuario" & vbCrLf & _
              "------", vbExclamation, "Permisos"
       Exit Sub
    End If
    
    With frmChat.sckClient
        .Close
        bolRecon = False
        .RemoteHost = txtServer.Text
        .RemotePort = CInt(txtPort.Text)
        .Connect
    End With
    
    Me.Hide
    frmChat.Show
    AddStatusMessage frmChat.rtbChat, RGB(128, 128, 128), "> Connecting to " & txtServer.Text & ":" & txtPort.Text & "..."
End Sub


Private Sub Form_Load()
   
   userName = Environ("USERNAME")
   UserDomain = Environ("USERDOMAIN")
   txtNickname.Text = LCase(userName)
   CargarConfiguracion
   
End Sub

Private Sub txtNickname_GotFocus()
   txtNickname.SelStart = 0
   txtNickname.SelLength = Len(txtNickname.Text)
End Sub


Private Sub txtPort_KeyPress(KeyAscii As Integer)
    'Number only.
    If Not IsNumeric(Chr$(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub


Private Sub CargarConfiguracion()
   Dim fs As New FileSystemObject
   Dim archivo As TextStream
   Dim linea As String
   
   Set archivo = fs.OpenTextFile(App.Path & "\chat.ini", ForReading)
   Do While archivo.AtEndOfStream = False
      linea = archivo.ReadLine
      If InStr(1, linea, "servidor") > 0 Then
         txtServer.Text = Replace(linea, "servidor:", "")
         gServidor = txtServer.Text
      End If
      If InStr(1, linea, "puerto") > 0 Then
         txtPort.Text = Replace(linea, "puerto:", "")
         gPuerto = txtPort.Text
      End If
   Loop
   archivo.Close
   Set fs = Nothing
   
End Sub

