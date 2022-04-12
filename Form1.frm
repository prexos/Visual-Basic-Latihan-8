VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   4800
      TabIndex        =   3
      Top             =   600
      Width           =   4815
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "POLITEKNIK NEGERI SEMARANG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   4
         Top             =   600
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GANTI JENIS FONT"
      Height          =   855
      Left            =   8160
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6000
      TabIndex        =   1
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GANTI PESAN"
      Height          =   855
      Left            =   3720
      TabIndex        =   0
      Top             =   3000
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Label1.Caption = InputBox("Silahkan:", "Ganti Pesan")
End Sub

Private Sub Command2_Click()
Dim GantiFont As Integer
GantiFont = MsgBox("Apakah ingin ganti font?", vbYesNoCancel, "Ganti Font")
 If GantiFont = vbYes Then
 Label1.FontName = "arial"
 Else
 Label1.FontName = "modern"
 End If
End Sub

Private Sub Exit_Click()
 Dim Respon As Integer
 Respon = MsgBox("Apakah Anda ingin keluar?", vbYesNo, "ALERT")
  If Respon = vbYes Then
   End
  End If
End Sub
