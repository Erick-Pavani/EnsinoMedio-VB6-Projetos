VERSION 5.00
Begin VB.Form frmInicio 
   BackColor       =   &H00FF0000&
   Caption         =   "Tutorial"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdJogar 
      Caption         =   "Ir para o jogo"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2640
      TabIndex        =   1
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label lblMensagem 
      BackColor       =   &H00FF0000&
      Caption         =   $"frmInicio.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdJogar_Click()
Me.Hide
frmSudoku.Show
End Sub
