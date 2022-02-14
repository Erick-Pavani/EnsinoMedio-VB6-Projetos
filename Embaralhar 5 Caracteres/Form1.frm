VERSION 5.00
Begin VB.Form frmEmbaralhar 
   Caption         =   "Embaralhar"
   ClientHeight    =   7215
   ClientLeft      =   4005
   ClientTop       =   2415
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   11805
   Begin VB.ListBox lstResultado 
      Height          =   2205
      Left            =   3000
      TabIndex        =   2
      Top             =   3720
      Width           =   6375
   End
   Begin VB.CommandButton cmdEmbaralhar 
      Caption         =   "Embaralhar"
      Height          =   1095
      Left            =   4200
      TabIndex        =   1
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox txtEntrada 
      Height          =   1095
      Left            =   1200
      TabIndex        =   0
      Text            =   "Digite a sequência (números ou letras) a ser embaralhados (5 caracteres) e em seguida clique no botão abaixo para concluir a ação!"
      Top             =   720
      Width           =   9615
   End
End
Attribute VB_Name = "frmEmbaralhar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Backup As String

Private Sub cmdEmbaralhar_Click()
Backup = txtEntrada.Text
A = Len(txtEntrada.Text)
If txtEntrada.Text = "" Or A < 5 Then
    Call MsgBox("Digite 5 caracteres para continuar!")
ElseIf txtEntrada.Text = "Digite a sequência (números ou letras) a ser embaralhados (5 caracteres) e em seguida clique no botão abaixo para concluir a ação!" Then
    Call MsgBox("Digite algo para continuar!")
Else
    Num1 = Left(txtEntrada.Text, 1)
    Num2 = Mid(txtEntrada.Text, 2, 1)
    Num3 = Mid(txtEntrada.Text, 3, 1)
    Num4 = Mid(txtEntrada.Text, 4, 1)
    Num5 = Right(txtEntrada.Text, 1)
    If Num1 = Num2 Or Num1 = Num3 Or Num1 = Num4 Or Num1 = Num5 Or Num2 = Num3 Or Num2 = Num4 Or Num2 = Num5 Or Num3 = Num4 Or Num3 = Num5 Or Num4 = Num5 Then
        Call MsgBox("A sequência não pode contêr caracteres repetidos!")
        txtEntrada.Text = ""
    Else
        X = 0
        Do While X <= 120
        lstResultado.AddItem (X)
        X = X + 1
        Loop
        Call MsgBox("Sua sequência foi embaralhada com sucesso!")
    End If
End If
End Sub

Private Sub txtEntrada_KeyPress(KeyAscii As Integer)
If txtEntrada.Text = "Digite a sequência (números ou letras) a ser embaralhados (5 caracteres) e em seguida clique no botão abaixo para concluir a ação!" Or txtEntrada.Text = Backup Then
    txtEntrada.Text = ""
    txtEntrada.MaxLength = 5
    lstResultado.Clear
End If
End Sub
