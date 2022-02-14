VERSION 5.00
Begin VB.Form frmEquilatero 
   BackColor       =   &H000000FF&
   Caption         =   "Equilátero"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVoltar1 
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
      Height          =   735
      Left            =   1680
      TabIndex        =   4
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label lblC2 
      BackColor       =   &H000000FF&
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
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblB2 
      BackColor       =   &H000000FF&
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
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblA2 
      BackColor       =   &H000000FF&
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
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   2400
      Width           =   255
   End
   Begin VB.Line linEquilatero3 
      BorderWidth     =   6
      X1              =   1800
      X2              =   4080
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line linEquilatero2 
      BorderWidth     =   6
      X1              =   3000
      X2              =   4080
      Y1              =   2280
      Y2              =   3120
   End
   Begin VB.Line linEquilatero1 
      BorderWidth     =   6
      X1              =   3000
      X2              =   1800
      Y1              =   2280
      Y2              =   3120
   End
   Begin VB.Label lblEquilatero 
      BackColor       =   &H000000FF&
      Caption         =   "Equilátero"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "frmEquilatero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVoltar1_Click()
Me.Hide
frmTriangulos.Show
frmTriangulos.txt1.Text = ""
frmTriangulos.txt2.Text = ""
frmTriangulos.txt3.Text = ""
End Sub
