VERSION 5.00
Begin VB.Form frmInicio 
   BackColor       =   &H00FF0000&
   Caption         =   "Jogo da Velha"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdComfirmar 
      Caption         =   "Comfirmar"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   2
      Top             =   4320
      Width           =   4215
   End
   Begin VB.TextBox txt2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblUser2 
      BackColor       =   &H00FF0000&
      Caption         =   "Usuário 2 "
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   5
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblUser1 
      BackColor       =   &H00FF0000&
      Caption         =   "Usuário 1"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblJogo 
      BackColor       =   &H00FF0000&
      Caption         =   "Jogo da Velha"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdComfirmar_Click()
If txt1.Text = "" Or txt2 = "" Or IsNumeric(txt1.Text) = True Or IsNumeric(txt2.Text) = True Or txt1.Text = txt2.Text Then
Call MsgBox("Por favor digite seu nome!")
Else
frmJogo.Show
Me.Hide
End If
frmJogo.lblPlayer1.Caption = txt1.Text
frmJogo.lblPlayer2.Caption = txt2.Text
frmJogo.lblPlacar1.Caption = "  0  "
frmJogo.lblPlacar2.Caption = "  0  "
Placar1 = 0
Placar2 = 0
Vez = 0
End Sub


