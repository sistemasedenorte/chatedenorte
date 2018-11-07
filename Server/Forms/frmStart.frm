VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada Servidor de Chat"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Iniciar »"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtPort 
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
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "1234"
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtNickname 
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Server"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   4560
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   4560
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Port:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nickname:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Username As String
Dim UserDomain As String


Private Sub cmdStart_Click()
    'Check input.
    txtNickname.Text = Trim$(Replace$(Replace$(txtNickname.Text, Chr$(2), ""), Chr$(4), ""))
    txtPort.Text = Trim$(txtPort.Text)
    
    If Len(txtNickname.Text) = 0 Then
        MsgBox "Please enter a nickname", vbCritical
        txtNickname.SetFocus
        Exit Sub
    ElseIf Len(txtPort.Text) = 0 Then
        MsgBox "Please enter a port", vbCritical
        txtPort.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txtPort.Text) Then
        MsgBox "Invalid port! Port must be numeric", vbCritical
        On Error Resume Next 'yea, yea...
        txtPort.SetFocus 'Can raise error.
        'Select all text.
        txtPort.SelStart = 0
        txtPort.SelLength = Len(txtPort.Text)
        Exit Sub
    End If
    
    'Code got this far, input is okay...
    'Close server first.
    ModChat.CloseServer
    'Now open server and hide this form.
    With frmChat
        .sckServer(0).Close
        .sckServer(0).LocalPort = CInt(txtPort.Text)
        .sckServer(0).Listen 'Opens the winsock control.
    End With
    
    ModServer.strMyNickname = txtNickname.Text
    frmChat.Show
    frmChat.lstUsers.Clear
    frmChat.lstUsers.AddItem strMyNickname
    AddStatusMessage frmChat.rtbChat, RGB(0, 128, 0), "> Server started on local IP (" & frmChat.sckServer(0).LocalIP & ":" & txtPort.Text & ")."
    frmChat.txtMsg.SetFocus
    Me.Hide
End Sub

Private Sub Form_Load()
    
    Username = Environ("USERNAME")
    UserDomain = Environ("USERDOMAIN")
    'txtNickname.Text = UserName
    
End Sub


Private Sub txtPort_KeyPress(KeyAscii As Integer)
    'Numbers only.
    If Not IsNumeric(Chr$(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub
