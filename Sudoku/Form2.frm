VERSION 5.00
Begin VB.Form frmGanhou 
   BackColor       =   &H0000FF00&
   Caption         =   "Ganhou"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNao 
      Caption         =   "Não"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      TabIndex        =   4
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdSim 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sim"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      TabIndex        =   3
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label lblJogarNovamente 
      BackColor       =   &H0000FF00&
      Caption         =   "Deseja jogar novamente?"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   3720
      Width           =   4935
   End
   Begin VB.Label lblVoceGanhou 
      BackColor       =   &H0000FF00&
      Caption         =   "Você terminou o Jogo"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label lblParabens 
      BackColor       =   &H0000FF00&
      Caption         =   "Parabéns!!!"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frmGanhou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNao_Click()
End
End Sub
Private Sub cmdSim_Click()
Me.Hide
frmSudoku.Show
End Sub
