VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt1 
      Height          =   1215
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   5055
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   1095
      Left            =   2880
      TabIndex        =   0
      Top             =   3840
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfirmar_Click()
A = txt1.Text
If txt1.Text = "" Then
    Call MsgBox("Digite algo para continuar!")
ElseIf IsNumeric(txt1.Text) Then
    Call MsgBox("Digite apenas letras!")
    txt1.Text = ""
ElseIf Len(txt1.Text) < 3 Then
    Call MsgBox("Digite mais de 3 carateres!")
Else
B = Len(txt1.Text)
D = 1
    Do While X < B
        C = Right(txt1.Text, 1)
        txt1.Text = Left(txt1.Text, B - D)
        Resultado = Resultado + C
        D = D + 1
        X = X + 1
    Loop
If UCase(Resultado) = UCase(A) Then
    Call MsgBox("É um palíndromo!")
    txt1.Text = ""
Else
    Call MsgBox("Não é um palíndromo!")
    txt1.Text = ""
End If
End If
End Sub
