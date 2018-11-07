VERSION 5.00
Object = "{546C0534-0DE2-457D-ACB3-531B0833BC86}#1.0#0"; "VB Splitter.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VBSplitter.Splitter Splitter 
      Height          =   5175
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9128
      FillContainer   =   0   'False
      Begin VB.Frame Frame 
         Height          =   5175
         Left            =   4545
         TabIndex        =   2
         Top             =   0
         Width           =   2190
      End
      Begin VB.TextBox Text 
         Height          =   5175
         Left            =   0
         TabIndex        =   1
         Text            =   "Text"
         Top             =   0
         Width           =   4485
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Splitter.Visible = True
End Sub
