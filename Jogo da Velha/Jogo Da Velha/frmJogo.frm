VERSION 5.00
Begin VB.Form frmJogo 
   BackColor       =   &H000000FF&
   Caption         =   "Jogo da Velha"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
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
      Left            =   7200
      TabIndex        =   22
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdFalse2 
      Caption         =   "False2"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   1
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdFalse1 
      Caption         =   "False1"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12240
      TabIndex        =   0
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdVoltar 
      Caption         =   "Trocar Jogadores"
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
      Left            =   7200
      TabIndex        =   19
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Frame fraJogoDaVelha 
      BackColor       =   &H000000FF&
      Caption         =   "Jogo da Velha"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   600
      TabIndex        =   7
      Top             =   2400
      Width           =   5295
      Begin VB.Line linJ8 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   6
         Visible         =   0   'False
         X1              =   720
         X2              =   4680
         Y1              =   600
         Y2              =   3120
      End
      Begin VB.Line linJ7 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   6
         Visible         =   0   'False
         X1              =   600
         X2              =   4800
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line linJ6 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   6
         Visible         =   0   'False
         X1              =   4680
         X2              =   600
         Y1              =   600
         Y2              =   3000
      End
      Begin VB.Line linJ5 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   6
         Visible         =   0   'False
         X1              =   2640
         X2              =   2640
         Y1              =   480
         Y2              =   3120
      End
      Begin VB.Line linJ4 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   6
         Visible         =   0   'False
         X1              =   4080
         X2              =   4080
         Y1              =   480
         Y2              =   3120
      End
      Begin VB.Line linJ3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   6
         Visible         =   0   'False
         X1              =   600
         X2              =   4800
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line linJ2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   6
         Visible         =   0   'False
         X1              =   600
         X2              =   4800
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line linJ1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   6
         Visible         =   0   'False
         X1              =   1200
         X2              =   1200
         Y1              =   480
         Y2              =   3120
      End
      Begin VB.Line lin4 
         BorderColor     =   &H00000000&
         BorderWidth     =   5
         X1              =   480
         X2              =   4800
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line lin3 
         BorderColor     =   &H00000000&
         BorderWidth     =   5
         X1              =   480
         X2              =   4800
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line lin2 
         BorderColor     =   &H00000000&
         BorderWidth     =   5
         X1              =   3360
         X2              =   3360
         Y1              =   360
         Y2              =   3240
      End
      Begin VB.Line lin1 
         BorderColor     =   &H00000000&
         BorderWidth     =   5
         X1              =   1920
         X2              =   1920
         Y1              =   360
         Y2              =   3240
      End
      Begin VB.Label lbl9 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   34.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3360
         TabIndex        =   18
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lbl8 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   34.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1920
         TabIndex        =   17
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lbl7 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   34.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   16
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lbl6 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   34.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3360
         TabIndex        =   15
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lbl5 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   34.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1920
         TabIndex        =   14
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lbl4 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   34.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lbl3 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   34.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3360
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lbl2 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   34.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1920
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lbl1 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   34.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H000000FF&
      Caption         =   "Selecione"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   6
      Top             =   360
      Width           =   3015
      Begin VB.OptionButton optO 
         BackColor       =   &H000000C0&
         Caption         =   " O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optX 
         BackColor       =   &H000000C0&
         Caption         =   " X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdReiniciar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reiniciar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   5
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Frame fraJogadores 
      BackColor       =   &H000000FF&
      Caption         =   "Jogadores"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   6480
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      Begin VB.Label lblPlacar2 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   21
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblPlacar1 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblPlayer2 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblPlayer1 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmJogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vez As Integer
Dim JogadorX As Integer
Dim JogadorO As Integer
Dim Placar1 As Integer
Dim Placar2 As Integer
Dim Ganhou As Boolean
'Feito por Erick! RA: 2140079-4 Todos os meus Trabalhos terão meu nome e RA no código!'
Private Sub cmdFalse2_GotFocus()
cmdFalse1.SetFocus
End Sub
Private Sub cmdReiniciar_Click()
lbl1.Caption = ""
lbl2.Caption = ""
lbl3.Caption = ""
lbl4.Caption = ""
lbl5.Caption = ""
lbl6.Caption = ""
lbl7.Caption = ""
lbl8.Caption = ""
lbl9.Caption = ""
linJ1.Visible = False
linJ2.Visible = False
linJ3.Visible = False
linJ4.Visible = False
linJ5.Visible = False
linJ6.Visible = False
linJ7.Visible = False
linJ8.Visible = False
optX.Value = False
optO.Value = False
fraOptions.Enabled = True
Vez = 0
End Sub
Private Sub cmdSair_Click()
If MsgBox("Deseja realmente sair", vbYesNo + vbQuestion, "Sair") = vbYes Then
End
Else
Cancel = True
End If
End Sub
Private Sub cmdVoltar_Click()
Me.Hide
frmInicio.Show
frmInicio.txt1.Text = ""
frmInicio.txt2.Text = ""
frmInicio.txt1.SetFocus
lbl1.Caption = ""
lbl2.Caption = ""
lbl3.Caption = ""
lbl4.Caption = ""
lbl5.Caption = ""
lbl6.Caption = ""
lbl7.Caption = ""
lbl8.Caption = ""
lbl9.Caption = ""
linJ1.Visible = False
linJ2.Visible = False
linJ3.Visible = False
linJ4.Visible = False
linJ5.Visible = False
linJ6.Visible = False
linJ7.Visible = False
linJ8.Visible = False
optX.Value = False
optO.Value = False
fraOptions.Enabled = True
Vez = 0
Placar1 = 0
Placar2 = 0
End Sub
Private Sub lbl1_Click()
If lbl1.Caption = "" Then
    If Vez = "1" Then
    lbl1.Caption = lbl1.Caption & "  X  "
    Vez = "2"
    Else
    If Vez = "2" Then
    lbl1.Caption = lbl1.Caption & "  O  "
    Vez = "1"
    End If
    End If
Else
Call MsgBox("Clique em um campo que ainda não foi preenchido!")
End If
        If lbl1.Caption = "  X  " And lbl2.Caption = "  X  " And lbl3.Caption = "  X  " Or lbl1.Caption = "  O  " And lbl2.Caption = "  O  " And lbl3.Caption = "  O  " Then
        linJ2.Visible = True
        lbl1.BackColor = &HFF00&
        lbl2.BackColor = &HFF00&
        lbl3.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl1.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl1.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl1.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl1.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ2.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl1.BackColor = &HC0&
            lbl2.BackColor = &HC0&
            lbl3.BackColor = &HC0&
        ElseIf lbl1.Caption = "  X  " And lbl4.Caption = "  X  " And lbl7.Caption = "  X  " Or lbl1.Caption = "  O  " And lbl4.Caption = "  O  " And lbl7.Caption = "  O  " Then
        linJ1.Visible = True
        lbl1.BackColor = &HFF00&
        lbl4.BackColor = &HFF00&
        lbl4.BackColor = &HFF00&
        If JogadorX = 1 And JogadorO = 2 Then
                If lbl1.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl1.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl1.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl1.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ1.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl1.BackColor = &HC0&
            lbl4.BackColor = &HC0&
            lbl7.BackColor = &HC0&
        ElseIf lbl1.Caption = "  X  " And lbl5.Caption = "  X  " And lbl9.Caption = "  X  " Or lbl1.Caption = "  O  " And lbl5.Caption = "  O  " And lbl9.Caption = "  O  " Then
        linJ8.Visible = True
        lbl1.BackColor = &HC0&
        lbl5.BackColor = &HC0&
        lbl9.BackColor = &HC0&
        If JogadorX = 1 And JogadorO = 2 Then
                If lbl1.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl1.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl1.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl1.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ8.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl1.BackColor = &HC0&
            lbl5.BackColor = &HC0&
            lbl9.BackColor = &HC0&
    ElseIf Not lbl1.Caption = "" And Not lbl2.Caption = "" And Not lbl3.Caption = "" And Not lbl4.Caption = "" And Not lbl5.Caption = "" And Not lbl6.Caption = "" And Not lbl7.Caption = "" And Not lbl8.Caption = "" And Not lbl9.Caption = "" And Ganhou = False Then
        lbl1.BackColor = &HFF0000
        lbl2.BackColor = &HFF0000
        lbl3.BackColor = &HFF0000
        lbl4.BackColor = &HFF0000
        lbl5.BackColor = &HFF0000
        lbl6.BackColor = &HFF0000
        lbl7.BackColor = &HFF0000
        lbl8.BackColor = &HFF0000
        lbl9.BackColor = &HFF0000
        Call MsgBox("Deu Velha! O jogo será reiniciado!")
        lbl1.Caption = ""
        lbl2.Caption = ""
        lbl3.Caption = ""
        lbl4.Caption = ""
        lbl5.Caption = ""
        lbl6.Caption = ""
        lbl7.Caption = ""
        lbl8.Caption = ""
        lbl9.Caption = ""
        optX.Value = False
        optO.Value = False
        fraOptions.Enabled = True
        Vez = 0
        lbl1.BackColor = &HC0&
        lbl2.BackColor = &HC0&
        lbl3.BackColor = &HC0&
        lbl4.BackColor = &HC0&
        lbl5.BackColor = &HC0&
        lbl6.BackColor = &HC0&
        lbl7.BackColor = &HC0&
        lbl8.BackColor = &HC0&
        lbl9.BackColor = &HC0&
        Call MsgBox("Nenhum dos jogadores ganharam ponto!")
        End If
End Sub
Private Sub lbl2_Click()
 If lbl2.Caption = "" Then
    If Vez = "1" Then
    lbl2.Caption = lbl2.Caption & "  X  "
    Vez = "2"
    Else
    If Vez = "2" Then
    lbl2.Caption = lbl2.Caption & "  O  "
    Vez = "1"
    End If
    End If
Else
Call MsgBox("Clique em um campo que ainda não foi preenchido!")
End If
        If lbl2.Caption = "  X  " And lbl5.Caption = "  X  " And lbl8.Caption = "  X  " Or lbl2.Caption = "  O  " And lbl5.Caption = "  O  " And lbl8.Caption = "  O  " Then
        linJ5.Visible = True
        lbl2.BackColor = &HFF00&
        lbl5.BackColor = &HFF00&
        lbl8.BackColor = &HFF00&
        If JogadorX = 1 And JogadorO = 2 Then
                If lbl2.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl2.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl2.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl2.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ5.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl2.BackColor = &HC0&
            lbl5.BackColor = &HC0&
            lbl8.BackColor = &HC0&
        ElseIf lbl1.Caption = "  X  " And lbl2.Caption = "  X  " And lbl3.Caption = "  X  " Or lbl1.Caption = "  O  " And lbl2.Caption = "  O  " And lbl3.Caption = "  O  " Then
        linJ2.Visible = True
        lbl1.BackColor = &HFF00&
        lbl2.BackColor = &HFF00&
        lbl3.BackColor = &HFF00&
        If JogadorX = 1 And JogadorO = 2 Then
                If lbl2.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl2.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl2.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl2.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ2.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl1.BackColor = &HC0&
            lbl2.BackColor = &HC0&
            lbl3.BackColor = &HC0&
        ElseIf Not lbl1.Caption = "" And Not lbl2.Caption = "" And Not lbl3.Caption = "" And Not lbl4.Caption = "" And Not lbl5.Caption = "" And Not lbl6.Caption = "" And Not lbl7.Caption = "" And Not lbl8.Caption = "" And Not lbl9.Caption = "" And Ganhou = False Then
        lbl1.BackColor = &HFF0000
        lbl2.BackColor = &HFF0000
        lbl3.BackColor = &HFF0000
        lbl4.BackColor = &HFF0000
        lbl5.BackColor = &HFF0000
        lbl6.BackColor = &HFF0000
        lbl7.BackColor = &HFF0000
        lbl8.BackColor = &HFF0000
        lbl9.BackColor = &HFF0000
        Call MsgBox("Deu Velha! O jogo será reiniciado!")
        lbl1.Caption = ""
        lbl2.Caption = ""
        lbl3.Caption = ""
        lbl4.Caption = ""
        lbl5.Caption = ""
        lbl6.Caption = ""
        lbl7.Caption = ""
        lbl8.Caption = ""
        lbl9.Caption = ""
        optX.Value = False
        optO.Value = False
        fraOptions.Enabled = True
        Vez = 0
        lbl1.BackColor = &HC0&
        lbl2.BackColor = &HC0&
        lbl3.BackColor = &HC0&
        lbl4.BackColor = &HC0&
        lbl5.BackColor = &HC0&
        lbl6.BackColor = &HC0&
        lbl7.BackColor = &HC0&
        lbl8.BackColor = &HC0&
        lbl9.BackColor = &HC0&
        Call MsgBox("Nenhum dos jogadores ganharam ponto!")
        End If
End Sub
Private Sub lbl3_Click()
 If lbl3.Caption = "" Then
    If Vez = "1" Then
    lbl3.Caption = lbl3.Caption & "  X  "
    Vez = "2"
    Else
    If Vez = "2" Then
    lbl3.Caption = lbl3.Caption & "  O  "
    Vez = "1"
    End If
    End If
Else
Call MsgBox("Clique em um campo que ainda não foi preenchido!")
End If
            If lbl3.Caption = "  X  " And lbl6.Caption = "  X  " And lbl9.Caption = "  X  " Or lbl3.Caption = "  O  " And lbl6.Caption = "  O  " And lbl9.Caption = "  O  " Then
            linJ4.Visible = True
            lbl3.BackColor = &HFF00&
            lbl6.BackColor = &HFF00&
            lbl9.BackColor = &HFF00&
                If JogadorX = 1 And JogadorO = 2 Then
                    If lbl3.Caption = "  X  " Then
                    MsgBox "O vencedor é : " & lblPlayer1.Caption
                    Placar1 = Placar1 + 1
                    Ganhou = True
                    ElseIf lbl3.Caption = "  O  " Then
                    MsgBox "O vencedor é : " & lblPlayer2.Caption
                    Placar2 = Placar2 + 1
                    Ganhou = True
                    End If
                ElseIf JogadorO = 1 And JogadorX = 2 Then
                    If lbl3.Caption = "  X  " Then
                    MsgBox "O vencedor é : " & lblPlayer2.Caption
                    Placar2 = Placar2 + 1
                    Ganhou = True
                    ElseIf lbl3.Caption = "  O  " Then
                    MsgBox "O vencedor é : " & lblPlayer1.Caption
                    Placar1 = Placar1 + 1
                    Ganhou = True
                    End If
                End If
                lbl1.Caption = ""
                lbl2.Caption = ""
                lbl3.Caption = ""
                lbl4.Caption = ""
                lbl5.Caption = ""
                lbl6.Caption = ""
                lbl7.Caption = ""
                lbl8.Caption = ""
                lbl9.Caption = ""
                linJ4.Visible = False
                optX.Value = False
                optO.Value = False
                fraOptions.Enabled = True
                Vez = 0
                lblPlacar1.Caption = Placar1
                lblPlacar2.Caption = Placar2
                Ganhou = False
                lbl3.BackColor = &HC0&
                lbl6.BackColor = &HC0&
                lbl9.BackColor = &HC0&
            ElseIf lbl3.Caption = "  X  " And lbl5.Caption = "  X  " And lbl7.Caption = "  X  " Or lbl3.Caption = "  O  " And lbl5.Caption = "  O  " And lbl7.Caption = "  O  " Then
            linJ6.Visible = True
            lbl3.BackColor = &HFF00&
            lbl5.BackColor = &HFF00&
            lbl7.BackColor = &HFF00&
                If JogadorX = 1 And JogadorO = 2 Then
                    If lbl3.Caption = "  X  " Then
                    MsgBox "O vencedor é : " & lblPlayer1.Caption
                    Placar1 = Placar1 + 1
                    Ganhou = True
                    ElseIf lbl3.Caption = "  O  " Then
                    MsgBox "O vencedor é : " & lblPlayer2.Caption
                    Placar2 = Placar2 + 1
                    Ganhou = True
                    End If
                ElseIf JogadorO = 1 And JogadorX = 2 Then
                    If lbl3.Caption = "  X  " Then
                    MsgBox "O vencedor é : " & lblPlayer2.Caption
                    Placar2 = Placar2 + 1
                    Ganhou = True
                    ElseIf lbl3.Caption = "  O  " Then
                    MsgBox "O vencedor é : " & lblPlayer1.Caption
                    Placar1 = Placar1 + 1
                    Ganhou = True
                    End If
                End If
                lbl1.Caption = ""
                lbl2.Caption = ""
                lbl3.Caption = ""
                lbl4.Caption = ""
                lbl5.Caption = ""
                lbl6.Caption = ""
                lbl7.Caption = ""
                lbl8.Caption = ""
                lbl9.Caption = ""
                linJ6.Visible = False
                optX.Value = False
                optO.Value = False
                fraOptions.Enabled = True
                Vez = 0
                lblPlacar1.Caption = Placar1
                lblPlacar2.Caption = Placar2
                Ganhou = False
                lbl3.BackColor = &HC0&
                lbl5.BackColor = &HC0&
                lbl7.BackColor = &HC0&
            ElseIf lbl1.Caption = "  X  " And lbl2.Caption = "  X  " And lbl3.Caption = "  X  " Or lbl1.Caption = "  O  " And lbl2.Caption = "  O  " And lbl3.Caption = "  O  " Then
            linJ2.Visible = True
            lbl1.BackColor = &HFF00&
            lbl2.BackColor = &HFF00&
            lbl3.BackColor = &HFF00&
                If JogadorX = 1 And JogadorO = 2 Then
                    If lbl3.Caption = "  X  " Then
                    MsgBox "O vencedor é : " & lblPlayer1.Caption
                    Placar1 = Placar1 + 1
                    Ganhou = True
                    ElseIf lbl3.Caption = "  O  " Then
                    MsgBox "O vencedor é : " & lblPlayer2.Caption
                    Placar2 = Placar2 + 1
                    Ganhou = True
                    End If
                ElseIf JogadorO = 1 And JogadorX = 2 Then
                    If lbl3.Caption = "  X  " Then
                    MsgBox "O vencedor é : " & lblPlayer2.Caption
                    Placar2 = Placar2 + 1
                    Ganhou = True
                    ElseIf lbl3.Caption = "  O  " Then
                    MsgBox "O vencedor é : " & lblPlayer1.Caption
                    Placar1 = Placar1 + 1
                    Ganhou = True
                    End If
                End If
                lbl1.Caption = ""
                lbl2.Caption = ""
                lbl3.Caption = ""
                lbl4.Caption = ""
                lbl5.Caption = ""
                lbl6.Caption = ""
                lbl7.Caption = ""
                lbl8.Caption = ""
                lbl9.Caption = ""
                linJ2.Visible = False
                optX.Value = False
                optO.Value = False
                fraOptions.Enabled = True
                Vez = 0
                lblPlacar1.Caption = Placar1
                lblPlacar2.Caption = Placar2
                Ganhou = False
                lbl1.BackColor = &HC0&
                lbl2.BackColor = &HC0&
                lbl3.BackColor = &HC0&
        ElseIf Not lbl1.Caption = "" And Not lbl2.Caption = "" And Not lbl3.Caption = "" And Not lbl4.Caption = "" And Not lbl5.Caption = "" And Not lbl6.Caption = "" And Not lbl7.Caption = "" And Not lbl8.Caption = "" And Not lbl9.Caption = "" And Ganhou = False Then
        lbl1.BackColor = &HFF0000
        lbl2.BackColor = &HFF0000
        lbl3.BackColor = &HFF0000
        lbl4.BackColor = &HFF0000
        lbl5.BackColor = &HFF0000
        lbl6.BackColor = &HFF0000
        lbl7.BackColor = &HFF0000
        lbl8.BackColor = &HFF0000
        lbl9.BackColor = &HFF0000
        Call MsgBox("Deu Velha! O jogo será reiniciado!")
        lbl1.Caption = ""
        lbl2.Caption = ""
        lbl3.Caption = ""
        lbl4.Caption = ""
        lbl5.Caption = ""
        lbl6.Caption = ""
        lbl7.Caption = ""
        lbl8.Caption = ""
        lbl9.Caption = ""
        optX.Value = False
        optO.Value = False
        fraOptions.Enabled = True
        Vez = 0
        lbl1.BackColor = &HC0&
        lbl2.BackColor = &HC0&
        lbl3.BackColor = &HC0&
        lbl4.BackColor = &HC0&
        lbl5.BackColor = &HC0&
        lbl6.BackColor = &HC0&
        lbl7.BackColor = &HC0&
        lbl8.BackColor = &HC0&
        lbl9.BackColor = &HC0&
        Call MsgBox("Nenhum dos jogadores ganharam ponto!")
        End If
End Sub
Private Sub lbl4_Click()
 If lbl4.Caption = "" Then
    If Vez = "1" Then
    lbl4.Caption = lbl4.Caption & "  X  "
    Vez = "2"
    Else
    If Vez = "2" Then
    lbl4.Caption = lbl4.Caption & "  O  "
    Vez = "1"
    End If
    End If
Else
Call MsgBox("Clique em um campo que ainda não foi preenchido!")
End If
        If lbl4.Caption = "  X  " And lbl5.Caption = "  X  " And lbl6.Caption = "  X  " Or lbl4.Caption = "  O  " And lbl5.Caption = "  O  " And lbl6.Caption = "  O  " Then
        linJ7.Visible = True
        lbl4.BackColor = &HFF00&
        lbl5.BackColor = &HFF00&
        lbl6.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl4.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl4.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl4.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl4.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ7.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl4.BackColor = &HC0&
            lbl5.BackColor = &HC0&
            lbl6.BackColor = &HC0&
        ElseIf lbl1.Caption = "  X  " And lbl4.Caption = "  X  " And lbl7.Caption = "  X  " Or lbl1.Caption = "  O  " And lbl4.Caption = "  O  " And lbl7.Caption = "  O  " Then
        linJ1.Visible = True
        lbl1.BackColor = &HFF00&
        lbl4.BackColor = &HFF00&
        lbl7.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl4.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl4.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl4.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl4.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
                lbl1.Caption = ""
                lbl2.Caption = ""
                lbl3.Caption = ""
                lbl4.Caption = ""
                lbl5.Caption = ""
                lbl6.Caption = ""
                lbl7.Caption = ""
                lbl8.Caption = ""
                lbl9.Caption = ""
                optX.Value = False
                optO.Value = False
                fraOptions.Enabled = True
                Vez = 0
                lblPlacar1.Caption = Placar1
                lblPlacar2.Caption = Placar2
                Ganhou = False
                linJ1.Visible = False
                lbl1.BackColor = &HC0&
                lbl4.BackColor = &HC0&
                lbl7.BackColor = &HC0&
        ElseIf Not lbl1.Caption = "" And Not lbl2.Caption = "" And Not lbl3.Caption = "" And Not lbl4.Caption = "" And Not lbl5.Caption = "" And Not lbl6.Caption = "" And Not lbl7.Caption = "" And Not lbl8.Caption = "" And Not lbl9.Caption = "" And Ganhou = False Then
        lbl1.BackColor = &HFF0000
        lbl2.BackColor = &HFF0000
        lbl3.BackColor = &HFF0000
        lbl4.BackColor = &HFF0000
        lbl5.BackColor = &HFF0000
        lbl6.BackColor = &HFF0000
        lbl7.BackColor = &HFF0000
        lbl8.BackColor = &HFF0000
        lbl9.BackColor = &HFF0000
        Call MsgBox("Deu Velha! O jogo será reiniciado!")
        lbl1.Caption = ""
        lbl2.Caption = ""
        lbl3.Caption = ""
        lbl4.Caption = ""
        lbl5.Caption = ""
        lbl6.Caption = ""
        lbl7.Caption = ""
        lbl8.Caption = ""
        lbl9.Caption = ""
        optX.Value = False
        optO.Value = False
        fraOptions.Enabled = True
        Vez = 0
        lbl1.BackColor = &HC0&
        lbl2.BackColor = &HC0&
        lbl3.BackColor = &HC0&
        lbl4.BackColor = &HC0&
        lbl5.BackColor = &HC0&
        lbl6.BackColor = &HC0&
        lbl7.BackColor = &HC0&
        lbl8.BackColor = &HC0&
        lbl9.BackColor = &HC0&
        Call MsgBox("Nenhum dos jogadores ganharam ponto!")
        End If
End Sub
Private Sub lbl5_Click()
 If lbl5.Caption = "" Then
    If Vez = "1" Then
    lbl5.Caption = lbl5.Caption & "  X  "
    Vez = "2"
    Else
    If Vez = "2" Then
    lbl5.Caption = lbl5.Caption & "  O  "
    Vez = "1"
    End If
    End If
Else
Call MsgBox("Clique em um campo que ainda não foi preenchido!")
End If
        If lbl1.Caption = "  X  " And lbl5.Caption = "  X  " And lbl9.Caption = "  X  " Or lbl1.Caption = "  O  " And lbl5.Caption = "  O  " And lbl9.Caption = "  O  " Then
        linJ8.Visible = True
        lbl1.BackColor = &HFF00&
        lbl5.BackColor = &HFF00&
        lbl9.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl5.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl5.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl5.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl5.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ8.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl1.BackColor = &HC0&
            lbl5.BackColor = &HC0&
            lbl9.BackColor = &HC0&
        ElseIf lbl2.Caption = "  X  " And lbl5.Caption = "  X  " And lbl8.Caption = "  X  " Or lbl2.Caption = "  O  " And lbl5.Caption = "  O  " And lbl8.Caption = "  O  " Then
        linJ5.Visible = True
        lbl2.BackColor = &HFF00&
        lbl5.BackColor = &HFF00&
        lbl8.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl5.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl5.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl5.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl5.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ5.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl2.BackColor = &HC0&
            lbl5.BackColor = &HC0&
            lbl8.BackColor = &HC0&
        ElseIf lbl3.Caption = "  X  " And lbl5.Caption = "  X  " And lbl7.Caption = "  X  " Or lbl3.Caption = "  O  " And lbl5.Caption = "  O  " And lbl7.Caption = "  O  " Then
        linJ6.Visible = True
        lbl3.BackColor = &HFF00&
        lbl5.BackColor = &HFF00&
        lbl7.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl5.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl5.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl5.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl5.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ6.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl3.BackColor = &HC0&
            lbl5.BackColor = &HC0&
            lbl7.BackColor = &HC0&
        ElseIf lbl4.Caption = "  X  " And lbl5.Caption = "  X  " And lbl6.Caption = "  X  " Or lbl4.Caption = "  O  " And lbl5.Caption = "  O  " And lbl6.Caption = "  O  " Then
        linJ7.Visible = True
        lbl4.BackColor = &HFF00&
        lbl5.BackColor = &HFF00&
        lbl6.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl5.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl5.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl5.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl5.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ7.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl4.BackColor = &HC0&
            lbl5.BackColor = &HC0&
            lbl6.BackColor = &HC0&
        ElseIf Not lbl1.Caption = "" And Not lbl2.Caption = "" And Not lbl3.Caption = "" And Not lbl4.Caption = "" And Not lbl5.Caption = "" And Not lbl6.Caption = "" And Not lbl7.Caption = "" And Not lbl8.Caption = "" And Not lbl9.Caption = "" And Ganhou = False Then
        lbl1.BackColor = &HFF0000
        lbl2.BackColor = &HFF0000
        lbl3.BackColor = &HFF0000
        lbl4.BackColor = &HFF0000
        lbl5.BackColor = &HFF0000
        lbl6.BackColor = &HFF0000
        lbl7.BackColor = &HFF0000
        lbl8.BackColor = &HFF0000
        lbl9.BackColor = &HFF0000
        Call MsgBox("Deu Velha! O jogo será reiniciado!")
        lbl1.Caption = ""
        lbl2.Caption = ""
        lbl3.Caption = ""
        lbl4.Caption = ""
        lbl5.Caption = ""
        lbl6.Caption = ""
        lbl7.Caption = ""
        lbl8.Caption = ""
        lbl9.Caption = ""
        optX.Value = False
        optO.Value = False
        fraOptions.Enabled = True
        Vez = 0
        lbl1.BackColor = &HC0&
        lbl2.BackColor = &HC0&
        lbl3.BackColor = &HC0&
        lbl4.BackColor = &HC0&
        lbl5.BackColor = &HC0&
        lbl6.BackColor = &HC0&
        lbl7.BackColor = &HC0&
        lbl8.BackColor = &HC0&
        lbl9.BackColor = &HC0&
        Call MsgBox("Nenhum dos jogadores ganharam ponto!")
        End If
End Sub
Private Sub lbl6_Click()
 If lbl6.Caption = "" Then
    If Vez = "1" Then
    lbl6.Caption = lbl6.Caption & "  X  "
    Vez = "2"
    Else
    If Vez = "2" Then
    lbl6.Caption = lbl6.Caption & "  O  "
    Vez = "1"
    End If
    End If
Else
Call MsgBox("Clique em um campo que ainda não foi preenchido!")
End If
        If lbl4.Caption = "  X  " And lbl5.Caption = "  X  " And lbl6.Caption = "  X  " Or lbl4.Caption = "  O  " And lbl5.Caption = "  O  " And lbl6.Caption = "  O  " Then
        linJ7.Visible = True
        lbl4.BackColor = &HFF00&
        lbl5.BackColor = &HFF00&
        lbl6.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl6.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl6.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl6.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl6.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ7.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl4.BackColor = &HC0&
            lbl5.BackColor = &HC0&
            lbl6.BackColor = &HC0&
        ElseIf lbl3.Caption = "  X  " And lbl6.Caption = "  X  " And lbl9.Caption = "  X  " Or lbl3.Caption = "  O  " And lbl6.Caption = "  O  " And lbl9.Caption = "  O  " Then
        linJ4.Visible = True
        lbl3.BackColor = &HFF00&
        lbl6.BackColor = &HFF00&
        lbl9.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl6.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl6.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl6.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl6.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ4.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl3.BackColor = &HC0&
            lbl6.BackColor = &HC0&
            lbl9.BackColor = &HC0&
        ElseIf Not lbl1.Caption = "" And Not lbl2.Caption = "" And Not lbl3.Caption = "" And Not lbl4.Caption = "" And Not lbl5.Caption = "" And Not lbl6.Caption = "" And Not lbl7.Caption = "" And Not lbl8.Caption = "" And Not lbl9.Caption = "" And Ganhou = False Then
        lbl1.BackColor = &HFF0000
        lbl2.BackColor = &HFF0000
        lbl3.BackColor = &HFF0000
        lbl4.BackColor = &HFF0000
        lbl5.BackColor = &HFF0000
        lbl6.BackColor = &HFF0000
        lbl7.BackColor = &HFF0000
        lbl8.BackColor = &HFF0000
        lbl9.BackColor = &HFF0000
        Call MsgBox("Deu Velha! O jogo será reiniciado!")
        lbl1.Caption = ""
        lbl2.Caption = ""
        lbl3.Caption = ""
        lbl4.Caption = ""
        lbl5.Caption = ""
        lbl6.Caption = ""
        lbl7.Caption = ""
        lbl8.Caption = ""
        lbl9.Caption = ""
        optX.Value = False
        optO.Value = False
        fraOptions.Enabled = True
        Vez = 0
        lbl1.BackColor = &HC0&
        lbl2.BackColor = &HC0&
        lbl3.BackColor = &HC0&
        lbl4.BackColor = &HC0&
        lbl5.BackColor = &HC0&
        lbl6.BackColor = &HC0&
        lbl7.BackColor = &HC0&
        lbl8.BackColor = &HC0&
        lbl9.BackColor = &HC0&
        Call MsgBox("Nenhum dos jogadores ganharam ponto!")
        End If
End Sub
Private Sub lbl7_Click()
 If lbl7.Caption = "" Then
    If Vez = "1" Then
    lbl7.Caption = lbl7.Caption & "  X  "
    Vez = "2"
    Else
    If Vez = "2" Then
    lbl7.Caption = lbl7.Caption & "  O  "
    Vez = "1"
    End If
    End If
Else
Call MsgBox("Clique em um campo que ainda não foi preenchido!")
End If
        If lbl1.Caption = "  X  " And lbl4.Caption = "  X  " And lbl7.Caption = "  X  " Or lbl1.Caption = "  O  " And lbl4.Caption = "  O  " And lbl7.Caption = "  O  " Then
        linJ1.Visible = True
        lbl1.BackColor = &HFF00&
        lbl4.BackColor = &HFF00&
        lbl7.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl7.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl7.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl7.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl7.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ1.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl1.BackColor = &HC0&
            lbl4.BackColor = &HC0&
            lbl7.BackColor = &HC0&
        ElseIf lbl3.Caption = "  X  " And lbl5.Caption = "  X  " And lbl7.Caption = "  X  " Or lbl3.Caption = "  O  " And lbl5.Caption = "  O  " And lbl7.Caption = "  O  " Then
        linJ6.Visible = True
        lbl3.BackColor = &HFF00&
        lbl5.BackColor = &HFF00&
        lbl7.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl7.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl7.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl7.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl7.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ6.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl3.BackColor = &HC0&
            lbl5.BackColor = &HC0&
            lbl7.BackColor = &HC0&
        ElseIf lbl7.Caption = "  X  " And lbl8.Caption = "  X  " And lbl9.Caption = "  X  " Or lbl7.Caption = "  O  " And lbl8.Caption = "  O  " And lbl9.Caption = "  O  " Then
        linJ3.Visible = True
        lbl7.BackColor = &HFF00&
        lbl8.BackColor = &HFF00&
        lbl9.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl7.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl7.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl7.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl7.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ3.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl7.BackColor = &HC0&
            lbl8.BackColor = &HC0&
            lbl9.BackColor = &HC0&
        ElseIf Not lbl1.Caption = "" And Not lbl2.Caption = "" And Not lbl3.Caption = "" And Not lbl4.Caption = "" And Not lbl5.Caption = "" And Not lbl6.Caption = "" And Not lbl7.Caption = "" And Not lbl8.Caption = "" And Not lbl9.Caption = "" And Ganhou = False Then
        lbl1.BackColor = &HFF0000
        lbl2.BackColor = &HFF0000
        lbl3.BackColor = &HFF0000
        lbl4.BackColor = &HFF0000
        lbl5.BackColor = &HFF0000
        lbl6.BackColor = &HFF0000
        lbl7.BackColor = &HFF0000
        lbl8.BackColor = &HFF0000
        lbl9.BackColor = &HFF0000
        Call MsgBox("Deu Velha! O jogo será reiniciado!")
        lbl1.Caption = ""
        lbl2.Caption = ""
        lbl3.Caption = ""
        lbl4.Caption = ""
        lbl5.Caption = ""
        lbl6.Caption = ""
        lbl7.Caption = ""
        lbl8.Caption = ""
        lbl9.Caption = ""
        optX.Value = False
        optO.Value = False
        fraOptions.Enabled = True
        Vez = 0
        lbl1.BackColor = &HC0&
        lbl2.BackColor = &HC0&
        lbl3.BackColor = &HC0&
        lbl4.BackColor = &HC0&
        lbl5.BackColor = &HC0&
        lbl6.BackColor = &HC0&
        lbl7.BackColor = &HC0&
        lbl8.BackColor = &HC0&
        lbl9.BackColor = &HC0&
        Call MsgBox("Nenhum dos jogadores ganharam ponto!")
        End If
End Sub
Private Sub lbl8_Click()
 If lbl8.Caption = "" Then
    If Vez = "1" Then
    lbl8.Caption = lbl8.Caption & "  X  "
    Vez = "2"
    Else
    If Vez = "2" Then
    lbl8.Caption = lbl8.Caption & "  O  "
    Vez = "1"
    End If
    End If
Else
Call MsgBox("Clique em um campo que ainda não foi preenchido!")
End If
        If lbl7.Caption = "  X  " And lbl8.Caption = "  X  " And lbl9.Caption = "  X  " Or lbl7.Caption = "  O  " And lbl8.Caption = "  O  " And lbl9.Caption = "  O  " Then
        linJ3.Visible = True
        lbl7.BackColor = &HFF00&
        lbl8.BackColor = &HFF00&
        lbl9.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl8.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl8.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl8.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl8.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ3.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl7.BackColor = &HC0&
            lbl8.BackColor = &HC0&
            lbl9.BackColor = &HC0&
        ElseIf lbl2.Caption = "  X  " And lbl5.Caption = "  X  " And lbl8.Caption = "  X  " Or lbl2.Caption = "  O  " And lbl5.Caption = "  O  " And lbl8.Caption = "  O  " Then
        linJ5.Visible = True
        lbl2.BackColor = &HFF00&
        lbl5.BackColor = &HFF00&
        lbl8.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl8.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl8.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl8.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl8.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ5.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl2.BackColor = &HC0&
            lbl5.BackColor = &HC0&
            lbl8.BackColor = &HC0&
        ElseIf Not lbl1.Caption = "" And Not lbl2.Caption = "" And Not lbl3.Caption = "" And Not lbl4.Caption = "" And Not lbl5.Caption = "" And Not lbl6.Caption = "" And Not lbl7.Caption = "" And Not lbl8.Caption = "" And Not lbl9.Caption = "" And Ganhou = False Then
        lbl1.BackColor = &HFF0000
        lbl2.BackColor = &HFF0000
        lbl3.BackColor = &HFF0000
        lbl4.BackColor = &HFF0000
        lbl5.BackColor = &HFF0000
        lbl6.BackColor = &HFF0000
        lbl7.BackColor = &HFF0000
        lbl8.BackColor = &HFF0000
        lbl9.BackColor = &HFF0000
        Call MsgBox("Deu Velha! O jogo será reiniciado!")
        lbl1.Caption = ""
        lbl2.Caption = ""
        lbl3.Caption = ""
        lbl4.Caption = ""
        lbl5.Caption = ""
        lbl6.Caption = ""
        lbl7.Caption = ""
        lbl8.Caption = ""
        lbl9.Caption = ""
        optX.Value = False
        optO.Value = False
        fraOptions.Enabled = True
        Vez = 0
        lbl1.BackColor = &HC0&
        lbl2.BackColor = &HC0&
        lbl3.BackColor = &HC0&
        lbl4.BackColor = &HC0&
        lbl5.BackColor = &HC0&
        lbl6.BackColor = &HC0&
        lbl7.BackColor = &HC0&
        lbl8.BackColor = &HC0&
        lbl9.BackColor = &HC0&
        Call MsgBox("Nenhum dos jogadores ganharam ponto!")
        End If
End Sub
Private Sub lbl9_Click()
 If lbl9.Caption = "" Then
    If Vez = "1" Then
    lbl9.Caption = lbl9.Caption & "  X  "
    Vez = "2"
    Else
    If Vez = "2" Then
    lbl9.Caption = lbl9.Caption & "  O  "
    Vez = "1"
    End If
    End If
Else
Call MsgBox("Clique em um campo que ainda não foi preenchido!")
End If
        If lbl3.Caption = "  X  " And lbl6.Caption = "  X  " And lbl9.Caption = "  X  " Or lbl3.Caption = "  O  " And lbl6.Caption = "  O  " And lbl9.Caption = "  O  " Then
        linJ4.Visible = True
        lbl3.BackColor = &HFF00&
        lbl6.BackColor = &HFF00&
        lbl9.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl9.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl9.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl9.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl9.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ4.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl3.BackColor = &HC0&
            lbl6.BackColor = &HC0&
            lbl9.BackColor = &HC0&
        ElseIf lbl7.Caption = "  X  " And lbl8.Caption = "  X  " And lbl9.Caption = "  X  " Or lbl7.Caption = "  O  " And lbl8.Caption = "  O  " And lbl9.Caption = "  O  " Then
        linJ3.Visible = True
        lbl7.BackColor = &HFF00&
        lbl8.BackColor = &HFF00&
        lbl9.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl9.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl9.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl9.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl9.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ3.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl7.BackColor = &HC0&
            lbl8.BackColor = &HC0&
            lbl9.BackColor = &HC0&
        ElseIf lbl1.Caption = "  X  " And lbl5.Caption = "  X  " And lbl9.Caption = "  X  " Or lbl1.Caption = "  O  " And lbl5.Caption = "  O  " And lbl9.Caption = "  O  " Then
        linJ8.Visible = True
        lbl1.BackColor = &HFF00&
        lbl5.BackColor = &HFF00&
        lbl9.BackColor = &HFF00&
            If JogadorX = 1 And JogadorO = 2 Then
                If lbl9.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                ElseIf lbl9.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                End If
            ElseIf JogadorO = 1 And JogadorX = 2 Then
                If lbl9.Caption = "  X  " Then
                MsgBox "O vencedor é : " & lblPlayer2.Caption
                Placar2 = Placar2 + 1
                Ganhou = True
                ElseIf lbl9.Caption = "  O  " Then
                MsgBox "O vencedor é : " & lblPlayer1.Caption
                Placar1 = Placar1 + 1
                Ganhou = True
                End If
            End If
            lbl1.Caption = ""
            lbl2.Caption = ""
            lbl3.Caption = ""
            lbl4.Caption = ""
            lbl5.Caption = ""
            lbl6.Caption = ""
            lbl7.Caption = ""
            lbl8.Caption = ""
            lbl9.Caption = ""
            linJ8.Visible = False
            optX.Value = False
            optO.Value = False
            fraOptions.Enabled = True
            Vez = 0
            lblPlacar1.Caption = Placar1
            lblPlacar2.Caption = Placar2
            Ganhou = False
            lbl1.BackColor = &HC0&
            lbl5.BackColor = &HC0&
            lbl9.BackColor = &HC0&
        ElseIf Not lbl1.Caption = "" And Not lbl2.Caption = "" And Not lbl3.Caption = "" And Not lbl4.Caption = "" And Not lbl5.Caption = "" And Not lbl6.Caption = "" And Not lbl7.Caption = "" And Not lbl8.Caption = "" And Not lbl9.Caption = "" And Ganhou = False Then
        Call MsgBox("Deu Velha! O jogo será reiniciado!")
        lbl1.Caption = ""
        lbl2.Caption = ""
        lbl3.Caption = ""
        lbl4.Caption = ""
        lbl5.Caption = ""
        lbl6.Caption = ""
        lbl7.Caption = ""
        lbl8.Caption = ""
        lbl9.Caption = ""
        optX.Value = False
        optO.Value = False
        fraOptions.Enabled = True
        Vez = 0
        lbl1.BackColor = &HC0&
        lbl2.BackColor = &HC0&
        lbl3.BackColor = &HC0&
        lbl4.BackColor = &HC0&
        lbl5.BackColor = &HC0&
        lbl6.BackColor = &HC0&
        lbl7.BackColor = &HC0&
        lbl8.BackColor = &HC0&
        lbl9.BackColor = &HC0&
        Call MsgBox("Nenhum dos jogadores ganharam ponto!")
        End If
End Sub
Private Sub optO_Click()
If optO.Value = True Then
JogadorO = 1
Vez = "2"
JogadorX = 2
End If
fraOptions.Enabled = False
End Sub
Private Sub optX_Click()
If optX.Value = True Then
JogadorX = 1
Vez = "1"
JogadorO = 2
End If
fraOptions.Enabled = False
End Sub
