VERSION 5.00
Begin VB.Form frmIsosceles 
   BackColor       =   &H0000FFFF&
   Caption         =   "Isósceles"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVoltar3 
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
      Left            =   1920
      TabIndex        =   0
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label lblC3 
      BackColor       =   &H0000FFFF&
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
      Left            =   3240
      TabIndex        =   4
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lblB3 
      BackColor       =   &H0000FFFF&
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
      Left            =   3840
      TabIndex        =   3
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblA3 
      BackColor       =   &H0000FFFF&
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
      Left            =   2640
      TabIndex        =   2
      Top             =   2040
      Width           =   255
   End
   Begin VB.Line linC3 
      BorderWidth     =   6
      X1              =   2520
      X2              =   4080
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line linB3 
      BorderWidth     =   6
      X1              =   3360
      X2              =   4080
      Y1              =   1560
      Y2              =   3360
   End
   Begin VB.Line linA3 
      BorderWidth     =   6
      X1              =   3360
      X2              =   2520
      Y1              =   1560
      Y2              =   3360
   End
   Begin VB.Label lblIsosceles 
      BackColor       =   &H0000FFFF&
      Caption         =   "Isósceles"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmIsosceles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVoltar3_Click()
Me.Hide
frmTriangulos.Show
frmTriangulos.txt1.Text = ""
frmTriangulos.txt2.Text = ""
frmTriangulos.txt3.Text = ""
End Sub
