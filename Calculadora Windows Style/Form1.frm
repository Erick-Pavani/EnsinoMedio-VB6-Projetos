VERSION 5.00
Begin VB.Form frmCalculadora 
   BackColor       =   &H00FF0000&
   Caption         =   "Calculadora"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFalse2 
      Caption         =   "False2"
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdFalse1 
      Caption         =   "False1"
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdVirgula 
      Caption         =   " ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   19
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdIgual 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   18
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   4680
      Width           =   5055
   End
   Begin VB.CommandButton cmdDivisao 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   16
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdMultiplicacao 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   15
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdSubtracao 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   14
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdicao 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   13
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   11
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblResultado 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmCalculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Num1 As Double
Dim Num2 As Double
Dim tam  As Single
Dim Sinal As Double
Private Sub cmd0_Click()
lblResultado.Caption = lblResultado.Caption & 0
cmdFalse1.SetFocus
End Sub
Private Sub cmd1_Click()
lblResultado.Caption = lblResultado.Caption & 1
cmdFalse1.SetFocus
End Sub
Private Sub cmd2_Click()
lblResultado.Caption = lblResultado.Caption & 2
cmdFalse1.SetFocus
End Sub
Private Sub cmd3_Click()
lblResultado.Caption = lblResultado.Caption & 3
cmdFalse1.SetFocus
End Sub
Private Sub cmd4_Click()
lblResultado.Caption = lblResultado.Caption & 4
cmdFalse1.SetFocus
End Sub
Private Sub cmd5_Click()
lblResultado.Caption = lblResultado.Caption & 5
cmdFalse1.SetFocus
End Sub
Private Sub cmd6_Click()
lblResultado.Caption = lblResultado.Caption & 6
cmdFalse1.SetFocus
End Sub
Private Sub cmd7_Click()
lblResultado.Caption = lblResultado.Caption & 7
cmdFalse1.SetFocus
End Sub
Private Sub cmd8_Click()
lblResultado.Caption = lblResultado.Caption & 8
cmdFalse1.SetFocus
End Sub
Private Sub cmd9_Click()
lblResultado.Caption = lblResultado.Caption & 9
cmdFalse1.SetFocus
End Sub
Private Sub cmdAdicao_Click()
Num1 = Val(lblResultado.Caption)
Sinal = "1"
lblResultado.Caption = ""
cmdIgual.SetFocus
End Sub
Private Sub cmdDivisao_Click()
Num1 = Val(lblResultado.Caption)
Sinal = "4"
lblResultado.Caption = ""
cmdIgual.SetFocus
End Sub
Private Sub cmdFalse1_KeyPress(KeyAscii As Integer)
If KeyAscii = 48 Then
    lblResultado.Caption = lblResultado.Caption & 0
ElseIf KeyAscii = 49 Then
    lblResultado.Caption = lblResultado.Caption & 1
ElseIf KeyAscii = 50 Then
    lblResultado.Caption = lblResultado.Caption & 2
ElseIf KeyAscii = 51 Then
    lblResultado.Caption = lblResultado.Caption & 3
ElseIf KeyAscii = 52 Then
    lblResultado.Caption = lblResultado.Caption & 4
ElseIf KeyAscii = 53 Then
    lblResultado.Caption = lblResultado.Caption & 5
ElseIf KeyAscii = 54 Then
    lblResultado.Caption = lblResultado.Caption & 6
ElseIf KeyAscii = 55 Then
    lblResultado.Caption = lblResultado.Caption & 7
ElseIf KeyAscii = 56 Then
    lblResultado.Caption = lblResultado.Caption & 8
ElseIf KeyAscii = 57 Then
    lblResultado.Caption = lblResultado.Caption & 9
ElseIf KeyAscii = 46 Then
    lblResultado.Caption = lblResultado.Caption & "."
ElseIf KeyAscii = 8 Then
    lblResultado.Caption = ""
ElseIf KeyAscii = 43 Then
    Num1 = Val(lblResultado.Caption)
    Sinal = "1"
    lblResultado.Caption = ""
    cmdIgual.SetFocus
ElseIf KeyAscii = 45 Then
    Num1 = Val(lblResultado.Caption)
    Sinal = "2"
    lblResultado.Caption = ""
    cmdIgual.SetFocus
ElseIf KeyAscii = 42 Then
    Num1 = Val(lblResultado.Caption)
    Sinal = "3"
    lblResultado.Caption = ""
    cmdIgual.SetFocus
ElseIf KeyAscii = 47 Then
    Num1 = Val(lblResultado.Caption)
    Sinal = "4"
    lblResultado.Caption = ""
    cmdIgual.SetFocus
End If
End Sub
Private Sub cmdFalse2_GotFocus()
cmdFalse1.SetFocus
End Sub
Private Sub cmdIgual_Click()
Num2 = Val(lblResultado.Caption)
If Sinal = "1" Then
    lblResultado = Num1 + Num2
ElseIf Sinal = "2" Then
    lblResultado = Num1 - Num2
ElseIf Sinal = "3" Then
    lblResultado = Num1 * Num2
ElseIf Sinal = "4" Then
    If Not Num2 = 0 Then
        lblResultado = Num1 / Num2
    Else
        Call MsgBox("Impossível Dividir Por Zero")
    End If
End If
cmdFalse1.SetFocus
End Sub
Private Sub cmdIgual_KeyPress(KeyAscii As Integer)
MsgBox (KeyAscii)
If KeyAscii = 48 Then
    lblResultado.Caption = lblResultado.Caption & 0
ElseIf KeyAscii = 49 Then
    lblResultado.Caption = lblResultado.Caption & 1
ElseIf KeyAscii = 50 Then
    lblResultado.Caption = lblResultado.Caption & 2
ElseIf KeyAscii = 51 Then
    lblResultado.Caption = lblResultado.Caption & 3
ElseIf KeyAscii = 52 Then
    lblResultado.Caption = lblResultado.Caption & 4
ElseIf KeyAscii = 53 Then
    lblResultado.Caption = lblResultado.Caption & 5
ElseIf KeyAscii = 54 Then
    lblResultado.Caption = lblResultado.Caption & 6
ElseIf KeyAscii = 55 Then
    lblResultado.Caption = lblResultado.Caption & 7
ElseIf KeyAscii = 56 Then
    lblResultado.Caption = lblResultado.Caption & 8
ElseIf KeyAscii = 57 Then
    lblResultado.Caption = lblResultado.Caption & 9
ElseIf KeyAscii = 46 Then
    lblResultado.Caption = lblResultado.Caption & "."
ElseIf KeyAscii = 8 Then
    MsgBox (" keyascii 8")
    tam = Len(lblResultado.Caption)
    lblResultado.Caption = Right(lblResultado.Caption, tam - 1)
ElseIf KeyAscii = 43 Then
    Num1 = Val(lblResultado.Caption) + Num1
    Sinal = "1"
    lblResultado.Caption = ""
    cmdIgual.SetFocus
ElseIf KeyAscii = 45 Then
    Num1 = Val(lblResultado.Caption) - Num1
    Sinal = "2"
    lblResultado.Caption = ""
    cmdIgual.SetFocus
ElseIf KeyAscii = 42 Then
    Num1 = Val(lblResultado.Caption) * Num1
    Sinal = "3"
    lblResultado.Caption = ""
    cmdIgual.SetFocus
ElseIf KeyAscii = 47 Then
    Num1 = Val(lblResultado.Caption) / Num1
    Sinal = "4"
    lblResultado.Caption = ""
    cmdIgual.SetFocus
End If
    End Sub
Private Sub cmdLimpar_Click()
lblResultado.Caption = ""
cmdFalse1.SetFocus
End Sub
Private Sub cmdMultiplicacao_Click()
Num1 = Val(lblResultado.Caption)
Sinal = "3"
lblResultado.Caption = ""
cmdIgual.SetFocus
End Sub
Private Sub cmdSubtracao_Click()
Num1 = Val(lblResultado.Caption)
Sinal = "2"
lblResultado.Caption = ""
cmdIgual.SetFocus
End Sub
Private Sub cmdVirgula_Click()
lblResultado.Caption = lblResultado.Caption & "."
cmdFalse1.SetFocus
End Sub
