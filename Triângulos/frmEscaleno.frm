VERSION 5.00
Begin VB.Form frmEscaleno 
   BackColor       =   &H00FF0000&
   Caption         =   "Escaleno"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVoltar2 
      Caption         =   "Voltar"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label lblC3 
      BackColor       =   &H00FF0000&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblB3 
      BackColor       =   &H00FF0000&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblA3 
      BackColor       =   &H00FF0000&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   375
   End
   Begin VB.Line linEscaleno1 
      BorderWidth     =   6
      X1              =   2640
      X2              =   1560
      Y1              =   1920
      Y2              =   3000
   End
   Begin VB.Line linEscaleno3 
      BorderWidth     =   6
      X1              =   1560
      X2              =   4560
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line linEscaleno2 
      BorderWidth     =   6
      X1              =   2640
      X2              =   4560
      Y1              =   1920
      Y2              =   3000
   End
   Begin VB.Label lblEscaleno 
      BackColor       =   &H00FF0000&
      Caption         =   "Escaleno"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmEscaleno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVoltar2_Click()
Me.Hide
frmTriangulos.Show
frmTriangulos.txt1.Text = ""
frmTriangulos.txt2.Text = ""
frmTriangulos.txt3.Text = ""
End Sub
