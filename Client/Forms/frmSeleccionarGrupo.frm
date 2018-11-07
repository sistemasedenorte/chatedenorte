VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmSeleccionarGrupo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enviar Mensaje a Grupo"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSeleccionarGrupo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtClave 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "clave"
      ToolTipText     =   "Indique la clave para envio de mensajes a grupos"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.ListBox lstGrupos 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5175
   End
   Begin ChamaleonButton.ChameleonBtn cmdAceptar 
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Aceptar"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSeleccionarGrupo.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdCancelar 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSeleccionarGrupo.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Seleccione el grupo al cual desea enviarle mensaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmSeleccionarGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public g_OK As Boolean
Public g_nombreGrupo As String


Private Sub cmdAceptar_Click()
    
    If Trim(txtClave.Text) <> Trim(strMyNickname) Then
      MsgBox "Indique la clave para enviar mensajes a grupos", vbExclamation, "Clave para Enviar"
      txtClave.SetFocus
      Exit Sub
    End If
    
    g_OK = True
    g_nombreGrupo = lstGrupos.List(lstGrupos.ListIndex)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
  g_OK = False
  Unload Me
End Sub

Private Sub Form_Load()
   g_OK = False
   g_nombreGrupo = ""
   CargarGrupos ""
End Sub

Private Sub lstGrupos_Click()
    cmdAceptar.Enabled = True
End Sub


'Me dice a cual  grupo pertenece una persona, id de red
Private Function CargarGrupos(idDeRed As String) As String
   Dim fs As New FileSystemObject
   Dim archivo As TextStream
   Dim linea, grupo As String
   
   'CargarListaPermisos
   lstGrupos.Clear
   For Each linea In gListaPermisos
      
      If InStr(1, linea, "<") > 0 And InStr(1, linea, "/") = 0 Then
         grupo = Replace(linea, "<", "")
         grupo = Replace(grupo, ">", "")
         lstGrupos.AddItem grupo
      End If

   Next
   
End Function

Private Sub lstGrupos_DblClick()
   cmdAceptar_Click
End Sub

Private Sub lstGrupos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      If cmdAceptar.Enabled = True Then cmdAceptar_Click
   End If
End Sub

Private Sub txtClave_GotFocus()
   txtClave.SelStart = 0
   txtClave.SelLength = Len(txtClave.Text)
End Sub
