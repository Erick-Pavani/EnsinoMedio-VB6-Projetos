VERSION 5.00
Begin VB.Form frmTriangulos 
   BackColor       =   &H0000FF00&
   Caption         =   "Triângulos"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txt2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txt3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   2400
      Width           =   2895
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
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
      Left            =   3120
      TabIndex        =   4
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   23.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   3
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label lblC 
      BackColor       =   &H0000FF00&
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
      Left            =   960
      TabIndex        =   10
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label lblB 
      BackColor       =   &H0000FF00&
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
      Left            =   1560
      TabIndex        =   9
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lblA 
      BackColor       =   &H0000FF00&
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
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   255
   End
   Begin VB.Line lin3 
      BorderWidth     =   9
      X1              =   240
      X2              =   1800
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line lin2 
      BorderWidth     =   9
      X1              =   1080
      X2              =   1800
      Y1              =   3960
      Y2              =   4800
   End
   Begin VB.Line lin1 
      BorderWidth     =   9
      X1              =   1080
      X2              =   240
      Y1              =   3960
      Y2              =   4800
   End
   Begin VB.Label lbl3 
      BackColor       =   &H0000FF00&
      Caption         =   "Lado C"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lbl2 
      BackColor       =   &H0000FF00&
      Caption         =   "Lado B"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lbl1 
      BackColor       =   &H0000FF00&
      Caption         =   "Lado A"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frmTriangulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lado1 As Double
Dim Lado2 As Double
Dim Lado3 As Double
'Feito por Erick! RA: 21340079-4 Todos os meus trabalhos terá meu nome e RA no código!'
Private Sub cmdConfirmar_Click()
Lado1 = Val(txt1.Text)
Lado2 = Val(txt2.Text)
Lado3 = Val(txt3.Text)
If Not IsNumeric(txt1.Text) Or Not IsNumeric(txt2.Text) Or Not IsNumeric(txt3.Text) Or Lado1 = 0 Or Lado2 = 0 Or Lado3 = 0 Then
Call MsgBox("Não é possível formar um triângulo com as medidas informadas!")
Call MsgBox("Por favor preencha todos os campos com números!")
Else
    If Lado1 < Lado2 + Lado3 And Lado2 < Lado1 + Lado3 And Lado3 < Lado1 + Lado2 Then
        If Lado1 > Lado2 And Lado1 > Lado3 And Lado2 <> Lado3 Then
            Me.Hide
            frmEscaleno.Show
            ElseIf Lado2 > Lado1 And Lado2 > Lado3 And Lado1 <> Lado3 Then
            Me.Hide
            frmEscaleno.Show
            ElseIf Lado3 > Lado2 And Lado3 > Lado1 And Lado2 <> Lado1 Then
            Me.Hide
            frmEscaleno.Show
                ElseIf Lado1 = Lado2 And Lado1 = Lado3 Then
                Me.Hide
                frmEquilatero.Show
                    ElseIf Lado1 = Lado2 And Lado1 <> Lado3 Then
                    Me.Hide
                    frmIsosceles.Show
                    ElseIf Lado1 = Lado3 And Lado1 <> Lado2 Then
                    Me.Hide
                    frmIsosceles.Show
                    ElseIf Lado2 = Lado3 And Lado2 <> Lado1 Then
                    Me.Hide
                    frmIsosceles.Show
        End If
    Else
    Call MsgBox("Não é possível formar um triângulo com as medidas informadas!")
    End If
End If
End Sub
Private Sub cmdLimpar_Click()
txt1.Text = ""
txt2.Text = ""
txt3.Text = ""
End Sub
