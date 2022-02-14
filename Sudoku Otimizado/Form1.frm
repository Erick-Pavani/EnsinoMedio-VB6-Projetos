VERSION 5.00
Begin VB.Form frmSudoku 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13845
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   13845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   10380
      TabIndex        =   83
      Top             =   6090
      Width           =   2415
   End
   Begin VB.ComboBox cmbNumeros 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      ItemData        =   "Form1.frx":0000
      Left            =   10440
      List            =   "Form1.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   82
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label lblLogo 
      BackColor       =   &H000000FF&
      Caption         =   "Sudoku"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   10200
      TabIndex        =   81
      Tag             =   "x"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   80
      Left            =   8760
      TabIndex        =   80
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   79
      Left            =   7920
      TabIndex        =   79
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   78
      Left            =   7080
      TabIndex        =   78
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   77
      Left            =   8760
      TabIndex        =   77
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   76
      Left            =   7920
      TabIndex        =   76
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   75
      Left            =   7080
      TabIndex        =   75
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   74
      Left            =   8760
      TabIndex        =   74
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   73
      Left            =   7920
      TabIndex        =   73
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   72
      Left            =   7080
      TabIndex        =   72
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   71
      Left            =   5640
      TabIndex        =   71
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   70
      Left            =   4800
      TabIndex        =   70
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   69
      Left            =   3960
      TabIndex        =   69
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   68
      Left            =   5640
      TabIndex        =   68
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   67
      Left            =   4800
      TabIndex        =   67
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   66
      Left            =   3960
      TabIndex        =   66
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   65
      Left            =   5640
      TabIndex        =   65
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   64
      Left            =   4800
      TabIndex        =   64
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   63
      Left            =   3960
      TabIndex        =   63
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   62
      Left            =   2520
      TabIndex        =   62
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   61
      Left            =   1680
      TabIndex        =   61
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   60
      Left            =   840
      TabIndex        =   60
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   59
      Left            =   2520
      TabIndex        =   59
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   58
      Left            =   1680
      TabIndex        =   58
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   57
      Left            =   840
      TabIndex        =   57
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   56
      Left            =   2520
      TabIndex        =   56
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   55
      Left            =   1680
      TabIndex        =   55
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   54
      Left            =   840
      TabIndex        =   54
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   53
      Left            =   8760
      TabIndex        =   53
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   52
      Left            =   7920
      TabIndex        =   52
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   51
      Left            =   7080
      TabIndex        =   51
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   50
      Left            =   8760
      TabIndex        =   50
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   49
      Left            =   7920
      TabIndex        =   49
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   48
      Left            =   7080
      TabIndex        =   48
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   47
      Left            =   8760
      TabIndex        =   47
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   46
      Left            =   7920
      TabIndex        =   46
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   45
      Left            =   7080
      TabIndex        =   45
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   44
      Left            =   5640
      TabIndex        =   44
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   43
      Left            =   4800
      TabIndex        =   43
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   42
      Left            =   3960
      TabIndex        =   42
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   41
      Left            =   5640
      TabIndex        =   41
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   40
      Left            =   4800
      TabIndex        =   40
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   39
      Left            =   3960
      TabIndex        =   39
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   38
      Left            =   5640
      TabIndex        =   38
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   37
      Left            =   4800
      TabIndex        =   37
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   36
      Left            =   3960
      TabIndex        =   36
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   35
      Left            =   2520
      TabIndex        =   35
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   34
      Left            =   1680
      TabIndex        =   34
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   33
      Left            =   840
      TabIndex        =   33
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   32
      Left            =   2520
      TabIndex        =   32
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   31
      Left            =   1680
      TabIndex        =   31
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   30
      Left            =   840
      TabIndex        =   30
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   29
      Left            =   2520
      TabIndex        =   29
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   28
      Left            =   1680
      TabIndex        =   28
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   27
      Left            =   840
      TabIndex        =   27
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   26
      Left            =   8760
      TabIndex        =   26
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   25
      Left            =   7920
      TabIndex        =   25
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   24
      Left            =   7080
      TabIndex        =   24
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   23
      Left            =   8760
      TabIndex        =   23
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   22
      Left            =   7920
      TabIndex        =   22
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   21
      Left            =   7080
      TabIndex        =   21
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   20
      Left            =   8760
      TabIndex        =   20
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   19
      Left            =   7920
      TabIndex        =   19
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   18
      Left            =   7080
      TabIndex        =   18
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   17
      Left            =   5640
      TabIndex        =   17
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   16
      Left            =   4800
      TabIndex        =   16
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   15
      Left            =   3960
      TabIndex        =   15
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   14
      Left            =   5640
      TabIndex        =   14
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   13
      Left            =   4800
      TabIndex        =   13
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   12
      Left            =   3960
      TabIndex        =   12
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   11
      Left            =   5640
      TabIndex        =   11
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   10
      Left            =   4800
      TabIndex        =   10
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   9
      Left            =   3960
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   8
      Left            =   2520
      TabIndex        =   8
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   7
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   6
      Left            =   840
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   5
      Left            =   2520
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   4
      Left            =   1680
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Tag             =   "A"
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "frmSudoku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Ganhou()
For x = 0 To 80
If lbl(x).ForeColor = &H0& Then
Next
Call MsgBox("Parabéns você ganhou o jogo!!")
lbl(x).Caption = "" And lbl(x).ForeColor = &H0&
End If
End Sub
Private Sub cmdLimpar_Click()
Call Limpar
End Sub
Private Sub lbl_Click(Index As Integer)
lbl.Item(Index).Caption = cmbNumeros.Text
    If lbl.Item(0).Caption = lbl.Item(1).Caption Or lbl.Item(0).Caption = lbl.Item(2).Caption Or lbl.Item(0).Caption = lbl.Item(9).Caption Or lbl.Item(0).Caption = lbl.Item(10).Caption Or lbl.Item(0).Caption = lbl.Item(11).Caption Or lbl.Item(0).Caption = lbl.Item(18).Caption Or lbl.Item(0).Caption = lbl.Item(19).Caption Or lbl.Item(0).Caption = lbl.Item(20).Caption Or lbl.Item(0).Caption = lbl.Item(3).Caption Or lbl.Item(0).Caption = lbl.Item(4).Caption Or lbl.Item(0).Caption = lbl.Item(5).Caption Or lbl.Item(0).Caption = lbl.Item(6).Caption Or lbl.Item(0).Caption = lbl.Item(7).Caption Or lbl.Item(0).Caption = lbl.Item(8).Caption Then
    lbl.Item(0).ForeColor = &HFF&
    ElseIf lbl.Item(9).Caption = lbl.Item(11).Caption Or lbl.Item(9).Caption = lbl.Item(0).Caption Or lbl.Item(9).Caption = lbl.Item(1).Caption Or lbl.Item(9).Caption = lbl.Item(2).Caption Or lbl.Item(9).Caption = lbl.Item(18).Caption Or lbl.Item(9).Caption = lbl.Item(19).Caption Or lbl.Item(9).Caption = lbl.Item(20).Caption Or lbl.Item(9).Caption = lbl.Item(12).Caption Or lbl.Item(9).Caption = lbl.Item(15).Caption Or lbl.Item(9).Caption = lbl.Item(36).Caption Or lbl.Item(9).Caption = lbl.Item(39).Caption Or lbl.Item(9).Caption = lbl.Item(42).Caption Or lbl.Item(9).Caption = lbl.Item(63).Caption Or lbl.Item(9).Caption = lbl.Item(66).Caption Or lbl.Item(9).Caption = lbl.Item(69).Caption Or lbl.Item(9).Caption = lbl.Item(13).Caption Or lbl.Item(9).Caption = lbl.Item(14).Caption Or lbl.Item(9).Caption = lbl.Item(16).Caption Or lbl.Item(9).Caption = lbl.Item(17).Caption Then
    lbl.Item(10).ForeColor = &HFF&
    ElseIf lbl.Item(10).Caption = lbl.Item(0).Caption Or lbl.Item(10).Caption = lbl.Item(1).Caption Or lbl.Item(10).Caption = lbl.Item(2).Caption Or lbl.Item(10).Caption = lbl.Item(10).Caption Or lbl.Item(10).Caption = lbl.Item(11).Caption Or lbl.Item(10).Caption = lbl.Item(18).Caption Or lbl.Item(10).Caption = lbl.Item(19).Caption Or lbl.Item(10).Caption = lbl.Item(20).Caption Or lbl.Item(10).Caption = lbl.Item(13).Caption Or lbl.Item(10).Caption = lbl.Item(16).Caption Or lbl.Item(10).Caption = lbl.Item(37).Caption Or lbl.Item(10).Caption = lbl.Item(40).Caption Or lbl.Item(10).Caption = lbl.Item(43).Caption Or lbl.Item(10).Caption = lbl.Item(64).Caption Or lbl.Item(10).Caption = lbl.Item(67).Caption Or lbl.Item(10).Caption = lbl.Item(70).Caption Or lbl.Item(10).Caption = lbl.Item(12).Caption Or lbl.Item(10).Caption = lbl.Item(14).Caption Or lbl.Item(10).Caption = lbl.Item(15).Caption Or lbl.Item(10).Caption = lbl.Item(17).Caption Then
    lbl.Item(10).ForeColor = &HFF&
    ElseIf lbl.Item(11).Caption = lbl.Item(0).Caption Or lbl.Item(11).Caption = lbl.Item(1).Caption Or lbl.Item(11).Caption = lbl.Item(2).Caption Or lbl.Item(11).Caption = lbl.Item(10).Caption Or lbl.Item(11).Caption = lbl.Item(10).Caption Or lbl.Item(11).Caption = lbl.Item(18).Caption Or lbl.Item(11).Caption = lbl.Item(19).Caption Or lbl.Item(11).Caption = lbl.Item(20).Caption Or lbl.Item(11).Caption = lbl.Item(14).Caption Or lbl.Item(11).Caption = lbl.Item(17).Caption Or lbl.Item(11).Caption = lbl.Item(38).Caption Or lbl.Item(11).Caption = lbl.Item(41).Caption Or lbl.Item(11).Caption = lbl.Item(44).Caption Or lbl.Item(11).Caption = lbl.Item(65).Caption Or lbl.Item(11).Caption = lbl.Item(68).Caption Or lbl.Item(11).Caption = lbl.Item(71).Caption Or lbl.Item(11).Caption = lbl.Item(12).Caption Or lbl.Item(11).Caption = lbl.Item(13).Caption Or lbl.Item(11).Caption = lbl.Item(15).Caption Or lbl.Item(11).Caption = lbl.Item(16).Caption Then
    lbl.Item(11).ForeColor = &HFF&
    ElseIf lbl.Item(12).Caption = lbl.Item(3).Caption Or lbl.Item(12).Caption = lbl.Item(4).Caption Or lbl.Item(12).Caption = lbl.Item(5).Caption Or lbl.Item(12).Caption = lbl.Item(13).Caption Or lbl.Item(12).Caption = lbl.Item(14).Caption Or lbl.Item(12).Caption = lbl.Item(21).Caption Or lbl.Item(12).Caption = lbl.Item(22).Caption Or lbl.Item(12).Caption = lbl.Item(23).Caption Or lbl.Item(12).Caption = lbl.Item(10).Caption Or lbl.Item(12).Caption = lbl.Item(15).Caption Or lbl.Item(12).Caption = lbl.Item(36).Caption Or lbl.Item(12).Caption = lbl.Item(39).Caption Or lbl.Item(12).Caption = lbl.Item(42).Caption Or lbl.Item(12).Caption = lbl.Item(63).Caption Or lbl.Item(12).Caption = lbl.Item(66).Caption Or lbl.Item(12).Caption = lbl.Item(69).Caption Or lbl.Item(12).Caption = lbl.Item(10).Caption Or lbl.Item(12).Caption = lbl.Item(11).Caption Or lbl.Item(12).Caption = lbl.Item(16).Caption Or lbl.Item(12).Caption = lbl.Item(17).Caption Then
    lbl.Item(12).ForeColor = &HFF&
    ElseIf lbl.Item(14).Caption = lbl.Item(10).Caption Or lbl.Item(14).Caption = lbl.Item(10).Caption Or lbl.Item(14).Caption = lbl.Item(11).Caption Or lbl.Item(14).Caption = lbl.Item(12).Caption Or lbl.Item(14).Caption = lbl.Item(13).Caption Or lbl.Item(14).Caption = lbl.Item(15).Caption Or lbl.Item(14).Caption = lbl.Item(16).Caption Or lbl.Item(14).Caption = lbl.Item(17).Caption Or lbl.Item(14).Caption = lbl.Item(3).Caption Or lbl.Item(14).Caption = lbl.Item(4).Caption Or lbl.Item(14).Caption = lbl.Item(5).Caption Or lbl.Item(14).Caption = lbl.Item(21).Caption Or lbl.Item(14).Caption = lbl.Item(22).Caption Or lbl.Item(14).Caption = lbl.Item(23).Caption Or lbl.Item(14).Caption = lbl.Item(38).Caption Or lbl.Item(14).Caption = lbl.Item(41).Caption Or lbl.Item(14).Caption = lbl.Item(44).Caption Or lbl.Item(14).Caption = lbl.Item(65).Caption Or lbl.Item(14).Caption = lbl.Item(68).Caption Or lbl.Item(14).Caption = lbl.Item(71).Caption Then
    lbl.Item(14).ForeColor = &HFF&
    ElseIf lbl.Item(15).Caption = lbl.Item(10).Caption Or lbl.Item(15).Caption = lbl.Item(10).Caption Or lbl.Item(15).Caption = lbl.Item(11).Caption Or lbl.Item(15).Caption = lbl.Item(12).Caption Or lbl.Item(15).Caption = lbl.Item(13).Caption Or lbl.Item(15).Caption = lbl.Item(14).Caption Or lbl.Item(15).Caption = lbl.Item(16).Caption Or lbl.Item(15).Caption = lbl.Item(17).Caption Or lbl.Item(15).Caption = lbl.Item(6).Caption Or lbl.Item(15).Caption = lbl.Item(7).Caption Or lbl.Item(15).Caption = lbl.Item(8).Caption Or lbl.Item(15).Caption = lbl.Item(24).Caption Or lbl.Item(15).Caption = lbl.Item(25).Caption Or lbl.Item(15).Caption = lbl.Item(26).Caption Or lbl.Item(15).Caption = lbl.Item(36).Caption Or lbl.Item(15).Caption = lbl.Item(39).Caption Or lbl.Item(15).Caption = lbl.Item(42).Caption Or lbl.Item(15).Caption = lbl.Item(63).Caption Or lbl.Item(15).Caption = lbl.Item(66).Caption Or lbl.Item(15).Caption = lbl.Item(69).Caption Then
    lbl.Item(15).ForeColor = &HFF&
    ElseIf lbl.Item(16).Caption = lbl.Item(6).Caption Or lbl.Item(16).Caption = lbl.Item(7).Caption Or lbl.Item(16).Caption = lbl.Item(8).Caption Or lbl.Item(16).Caption = lbl.Item(24).Caption Or lbl.Item(16).Caption = lbl.Item(25).Caption Or lbl.Item(16).Caption = lbl.Item(26).Caption Or lbl.Item(16).Caption = lbl.Item(10).Caption Or lbl.Item(16).Caption = lbl.Item(10).Caption Or lbl.Item(16).Caption = lbl.Item(11).Caption Or lbl.Item(16).Caption = lbl.Item(12).Caption Or lbl.Item(16).Caption = lbl.Item(13).Caption Or lbl.Item(16).Caption = lbl.Item(14).Caption Or lbl.Item(16).Caption = lbl.Item(15).Caption Or lbl.Item(16).Caption = lbl.Item(17).Caption Or lbl.Item(16).Caption = lbl.Item(37).Caption Or lbl.Item(16).Caption = lbl.Item(40).Caption Or lbl.Item(16).Caption = lbl.Item(43).Caption Or lbl.Item(16).Caption = lbl.Item(64).Caption Or lbl.Item(16).Caption = lbl.Item(67).Caption Or lbl.Item(16).Caption = lbl.Item(70).Caption Then
    lbl.Item(16).ForeColor = &HFF&
    ElseIf lbl.Item(17).Caption = lbl.Item(6).Caption Or lbl.Item(17).Caption = lbl.Item(7).Caption Or lbl.Item(17).Caption = lbl.Item(8).Caption Or lbl.Item(17).Caption = lbl.Item(24).Caption Or lbl.Item(17).Caption = lbl.Item(25).Caption Or lbl.Item(17).Caption = lbl.Item(26).Caption Or lbl.Item(17).Caption = lbl.Item(10).Caption Or lbl.Item(17).Caption = lbl.Item(10).Caption Or lbl.Item(17).Caption = lbl.Item(11).Caption Or lbl.Item(17).Caption = lbl.Item(12).Caption Or lbl.Item(17).Caption = lbl.Item(13).Caption Or lbl.Item(17).Caption = lbl.Item(14).Caption Or lbl.Item(17).Caption = lbl.Item(15).Caption Or lbl.Item(17).Caption = lbl.Item(16).Caption Or lbl.Item(17).Caption = lbl.Item(38).Caption Or lbl.Item(17).Caption = lbl.Item(41).Caption Or lbl.Item(17).Caption = lbl.Item(44).Caption Or lbl.Item(17).Caption = lbl.Item(65).Caption Or lbl.Item(17).Caption = lbl.Item(68).Caption Or lbl.Item(17).Caption = lbl.Item(71).Caption Then
    lbl.Item(17).ForeColor = &HFF&
    ElseIf lbl.Item(18).Caption = lbl.Item(0).Caption Or lbl.Item(18).Caption = lbl.Item(1).Caption Or lbl.Item(18).Caption = lbl.Item(2).Caption Or lbl.Item(18).Caption = lbl.Item(10).Caption Or lbl.Item(18).Caption = lbl.Item(10).Caption Or lbl.Item(18).Caption = lbl.Item(11).Caption Or lbl.Item(18).Caption = lbl.Item(19).Caption Or lbl.Item(18).Caption = lbl.Item(20).Caption Or lbl.Item(18).Caption = lbl.Item(21).Caption Or lbl.Item(18).Caption = lbl.Item(22).Caption Or lbl.Item(18).Caption = lbl.Item(23).Caption Or lbl.Item(18).Caption = lbl.Item(24).Caption Or lbl.Item(18).Caption = lbl.Item(25).Caption Or lbl.Item(18).Caption = lbl.Item(26).Caption Or lbl.Item(18).Caption = lbl.Item(45).Caption Or lbl.Item(18).Caption = lbl.Item(48).Caption Or lbl.Item(18).Caption = lbl.Item(51).Caption Or lbl.Item(18).Caption = lbl.Item(72).Caption Or lbl.Item(18).Caption = lbl.Item(75).Caption Or lbl.Item(18).Caption = lbl.Item(78).Caption Then
    lbl.Item(18).ForeColor = &HFF&
    ElseIf lbl.Item(1).Caption = lbl.Item(10).Caption Or lbl.Item(1).Caption = lbl.Item(10).Caption Or lbl.Item(1).Caption = lbl.Item(11).Caption Or lbl.Item(1).Caption = lbl.Item(18).Caption Or lbl.Item(1).Caption = lbl.Item(19).Caption Or lbl.Item(1).Caption = lbl.Item(20).Caption Or lbl.Item(1).Caption = lbl.Item(0).Caption Or lbl.Item(1).Caption = lbl.Item(2).Caption Or lbl.Item(1).Caption = lbl.Item(3).Caption Or lbl.Item(1).Caption = lbl.Item(4).Caption Or lbl.Item(1).Caption = lbl.Item(5).Caption Or lbl.Item(1).Caption = lbl.Item(6).Caption Or lbl.Item(1).Caption = lbl.Item(7).Caption Or lbl.Item(1).Caption = lbl.Item(8).Caption Or lbl.Item(1).Caption = lbl.Item(28).Caption Or lbl.Item(1).Caption = lbl.Item(31).Caption Or lbl.Item(1).Caption = lbl.Item(34).Caption Or lbl.Item(1).Caption = lbl.Item(55).Caption Or lbl.Item(1).Caption = lbl.Item(58).Caption Or lbl.Item(1).Caption = lbl.Item(61).Caption Then
    lbl.Item(1).ForeColor = &HFF&
    ElseIf lbl.Item(19).Caption = lbl.Item(0).Caption Or lbl.Item(19).Caption = lbl.Item(1).Caption Or lbl.Item(19).Caption = lbl.Item(2).Caption Or lbl.Item(19).Caption = lbl.Item(10).Caption Or lbl.Item(19).Caption = lbl.Item(10).Caption Or lbl.Item(19).Caption = lbl.Item(11).Caption Or lbl.Item(19).Caption = lbl.Item(18).Caption Or lbl.Item(19).Caption = lbl.Item(20).Caption Or lbl.Item(19).Caption = lbl.Item(21).Caption Or lbl.Item(19).Caption = lbl.Item(22).Caption Or lbl.Item(19).Caption = lbl.Item(23).Caption Or lbl.Item(19).Caption = lbl.Item(24).Caption Or lbl.Item(19).Caption = lbl.Item(25).Caption Or lbl.Item(19).Caption = lbl.Item(26).Caption Or lbl.Item(19).Caption = lbl.Item(46).Caption Or lbl.Item(19).Caption = lbl.Item(49).Caption Or lbl.Item(19).Caption = lbl.Item(52).Caption Or lbl.Item(19).Caption = lbl.Item(73).Caption Or lbl.Item(19).Caption = lbl.Item(76).Caption Or lbl.Item(19).Caption = lbl.Item(79).Caption Then
    lbl.Item(19).ForeColor = &HFF&
    ElseIf lbl.Item(20).Caption = lbl.Item(0).Caption Or lbl.Item(20).Caption = lbl.Item(1).Caption Or lbl.Item(20).Caption = lbl.Item(2).Caption Or lbl.Item(20).Caption = lbl.Item(10).Caption Or lbl.Item(20).Caption = lbl.Item(10).Caption Or lbl.Item(20).Caption = lbl.Item(11).Caption Or lbl.Item(20).Caption = lbl.Item(18).Caption Or lbl.Item(20).Caption = lbl.Item(19).Caption Or lbl.Item(20).Caption = lbl.Item(21).Caption Or lbl.Item(20).Caption = lbl.Item(22).Caption Or lbl.Item(20).Caption = lbl.Item(23).Caption Or lbl.Item(20).Caption = lbl.Item(24).Caption Or lbl.Item(20).Caption = lbl.Item(25).Caption Or lbl.Item(20).Caption = lbl.Item(26).Caption Or lbl.Item(20).Caption = lbl.Item(47).Caption Or lbl.Item(20).Caption = lbl.Item(50).Caption Or lbl.Item(20).Caption = lbl.Item(53).Caption Or lbl.Item(20).Caption = lbl.Item(74).Caption Or lbl.Item(20).Caption = lbl.Item(77).Caption Or lbl.Item(20).Caption = lbl.Item(80).Caption Then
    lbl.Item(20).ForeColor = &HFF&
    ElseIf lbl.Item(21).Caption = lbl.Item(3).Caption Or lbl.Item(21).Caption = lbl.Item(4).Caption Or lbl.Item(21).Caption = lbl.Item(5).Caption Or lbl.Item(21).Caption = lbl.Item(12).Caption Or lbl.Item(21).Caption = lbl.Item(13).Caption Or lbl.Item(21).Caption = lbl.Item(14).Caption Or lbl.Item(21).Caption = lbl.Item(18).Caption Or lbl.Item(21).Caption = lbl.Item(19).Caption Or lbl.Item(21).Caption = lbl.Item(20).Caption Or lbl.Item(21).Caption = lbl.Item(22).Caption Or lbl.Item(21).Caption = lbl.Item(23).Caption Or lbl.Item(21).Caption = lbl.Item(24).Caption Or lbl.Item(21).Caption = lbl.Item(25).Caption Or lbl.Item(21).Caption = lbl.Item(26).Caption Or lbl.Item(21).Caption = lbl.Item(45).Caption Or lbl.Item(21).Caption = lbl.Item(48).Caption Or lbl.Item(21).Caption = lbl.Item(51).Caption Or lbl.Item(21).Caption = lbl.Item(72).Caption Or lbl.Item(21).Caption = lbl.Item(75).Caption Or lbl.Item(21).Caption = lbl.Item(78).Caption Then
    lbl.Item(21).ForeColor = &HFF&
    ElseIf lbl.Item(22).Caption = lbl.Item(3).Caption Or lbl.Item(22).Caption = lbl.Item(4).Caption Or lbl.Item(22).Caption = lbl.Item(5).Caption Or lbl.Item(22).Caption = lbl.Item(12).Caption Or lbl.Item(22).Caption = lbl.Item(13).Caption Or lbl.Item(22).Caption = lbl.Item(14).Caption Or lbl.Item(22).Caption = lbl.Item(18).Caption Or lbl.Item(22).Caption = lbl.Item(19).Caption Or lbl.Item(22).Caption = lbl.Item(20).Caption Or lbl.Item(22).Caption = lbl.Item(21).Caption Or lbl.Item(22).Caption = lbl.Item(23).Caption Or lbl.Item(22).Caption = lbl.Item(24).Caption Or lbl.Item(22).Caption = lbl.Item(25).Caption Or lbl.Item(22).Caption = lbl.Item(26).Caption Or lbl.Item(22).Caption = lbl.Item(46).Caption Or lbl.Item(22).Caption = lbl.Item(49).Caption Or lbl.Item(22).Caption = lbl.Item(52).Caption Or lbl.Item(22).Caption = lbl.Item(73).Caption Or lbl.Item(22).Caption = lbl.Item(76).Caption Or lbl.Item(22).Caption = lbl.Item(79).Caption Then
    lbl.Item(22).ForeColor = &HFF&
    ElseIf lbl.Item(23).Caption = lbl.Item(3).Caption Or lbl.Item(23).Caption = lbl.Item(4).Caption Or lbl.Item(23).Caption = lbl.Item(5).Caption Or lbl.Item(23).Caption = lbl.Item(12).Caption Or lbl.Item(23).Caption = lbl.Item(13).Caption Or lbl.Item(23).Caption = lbl.Item(14).Caption Or lbl.Item(23).Caption = lbl.Item(18).Caption Or lbl.Item(23).Caption = lbl.Item(19).Caption Or lbl.Item(23).Caption = lbl.Item(20).Caption Or lbl.Item(23).Caption = lbl.Item(21).Caption Or lbl.Item(23).Caption = lbl.Item(22).Caption Or lbl.Item(23).Caption = lbl.Item(24).Caption Or lbl.Item(23).Caption = lbl.Item(25).Caption Or lbl.Item(23).Caption = lbl.Item(26).Caption Or lbl.Item(23).Caption = lbl.Item(47).Caption Or lbl.Item(23).Caption = lbl.Item(50).Caption Or lbl.Item(23).Caption = lbl.Item(53).Caption Or lbl.Item(23).Caption = lbl.Item(74).Caption Or lbl.Item(23).Caption = lbl.Item(77).Caption Or lbl.Item(23).Caption = lbl.Item(80).Caption Then
    lbl.Item(23).ForeColor = &HFF&
    ElseIf lbl.Item(24).Caption = lbl.Item(6).Caption Or lbl.Item(24).Caption = lbl.Item(7).Caption Or lbl.Item(24).Caption = lbl.Item(8).Caption Or lbl.Item(24).Caption = lbl.Item(15).Caption Or lbl.Item(24).Caption = lbl.Item(16).Caption Or lbl.Item(24).Caption = lbl.Item(17).Caption Or lbl.Item(24).Caption = lbl.Item(18).Caption Or lbl.Item(24).Caption = lbl.Item(19).Caption Or lbl.Item(24).Caption = lbl.Item(20).Caption Or lbl.Item(24).Caption = lbl.Item(21).Caption Or lbl.Item(24).Caption = lbl.Item(22).Caption Or lbl.Item(24).Caption = lbl.Item(23).Caption Or lbl.Item(24).Caption = lbl.Item(25).Caption Or lbl.Item(24).Caption = lbl.Item(26).Caption Or lbl.Item(24).Caption = lbl.Item(45).Caption Or lbl.Item(24).Caption = lbl.Item(48).Caption Or lbl.Item(24).Caption = lbl.Item(51).Caption Or lbl.Item(24).Caption = lbl.Item(72).Caption Or lbl.Item(24).Caption = lbl.Item(75).Caption Or lbl.Item(24).Caption = lbl.Item(78).Caption Then
    lbl.Item(24).ForeColor = &HFF&
    ElseIf lbl.Item(25).Caption = lbl.Item(6).Caption Or lbl.Item(25).Caption = lbl.Item(7).Caption Or lbl.Item(25).Caption = lbl.Item(8).Caption Or lbl.Item(25).Caption = lbl.Item(15).Caption Or lbl.Item(25).Caption = lbl.Item(16).Caption Or lbl.Item(25).Caption = lbl.Item(17).Caption Or lbl.Item(25).Caption = lbl.Item(18).Caption Or lbl.Item(25).Caption = lbl.Item(19).Caption Or lbl.Item(25).Caption = lbl.Item(20).Caption Or lbl.Item(25).Caption = lbl.Item(21).Caption Or lbl.Item(25).Caption = lbl.Item(22).Caption Or lbl.Item(25).Caption = lbl.Item(23).Caption Or lbl.Item(25).Caption = lbl.Item(24).Caption Or lbl.Item(25).Caption = lbl.Item(26).Caption Or lbl.Item(25).Caption = lbl.Item(46).Caption Or lbl.Item(25).Caption = lbl.Item(49).Caption Or lbl.Item(25).Caption = lbl.Item(52).Caption Or lbl.Item(25).Caption = lbl.Item(73).Caption Or lbl.Item(25).Caption = lbl.Item(76).Caption Or lbl.Item(25).Caption = lbl.Item(79).Caption Then
    lbl.Item(25).ForeColor = &HFF&
    ElseIf lbl.Item(27).Caption = lbl.Item(0).Caption Or lbl.Item(27).Caption = lbl.Item(3).Caption Or lbl.Item(27).Caption = lbl.Item(6).Caption Or lbl.Item(27).Caption = lbl.Item(54).Caption Or lbl.Item(27).Caption = lbl.Item(57).Caption Or lbl.Item(27).Caption = lbl.Item(60).Caption Or lbl.Item(27).Caption = lbl.Item(28).Caption Or lbl.Item(27).Caption = lbl.Item(29).Caption Or lbl.Item(27).Caption = lbl.Item(30).Caption Or lbl.Item(27).Caption = lbl.Item(31).Caption Or lbl.Item(27).Caption = lbl.Item(32).Caption Or lbl.Item(27).Caption = lbl.Item(33).Caption Or lbl.Item(27).Caption = lbl.Item(34).Caption Or lbl.Item(27).Caption = lbl.Item(35).Caption Or lbl.Item(27).Caption = lbl.Item(36).Caption Or lbl.Item(27).Caption = lbl.Item(37).Caption Or lbl.Item(27).Caption = lbl.Item(38).Caption Or lbl.Item(27).Caption = lbl.Item(45).Caption Or lbl.Item(27).Caption = lbl.Item(46).Caption Or lbl.Item(27).Caption = lbl.Item(47).Caption Then
    lbl.Item(27).ForeColor = &HFF&
    ElseIf lbl.Item(28).Caption = lbl.Item(1).Caption Or lbl.Item(28).Caption = lbl.Item(4).Caption Or lbl.Item(28).Caption = lbl.Item(7).Caption Or lbl.Item(28).Caption = lbl.Item(55).Caption Or lbl.Item(28).Caption = lbl.Item(58).Caption Or lbl.Item(28).Caption = lbl.Item(61).Caption Or lbl.Item(28).Caption = lbl.Item(27).Caption Or lbl.Item(28).Caption = lbl.Item(29).Caption Or lbl.Item(28).Caption = lbl.Item(30).Caption Or lbl.Item(28).Caption = lbl.Item(31).Caption Or lbl.Item(28).Caption = lbl.Item(32).Caption Or lbl.Item(28).Caption = lbl.Item(33).Caption Or lbl.Item(28).Caption = lbl.Item(34).Caption Or lbl.Item(28).Caption = lbl.Item(35).Caption Or lbl.Item(28).Caption = lbl.Item(36).Caption Or lbl.Item(28).Caption = lbl.Item(37).Caption Or lbl.Item(28).Caption = lbl.Item(38).Caption Or lbl.Item(28).Caption = lbl.Item(45).Caption Or lbl.Item(28).Caption = lbl.Item(46).Caption Or lbl.Item(28).Caption = lbl.Item(47).Caption Then
    lbl.Item(28).ForeColor = &HFF&
    ElseIf lbl.Item(2).Caption = lbl.Item(10).Caption Or lbl.Item(2).Caption = lbl.Item(10).Caption Or lbl.Item(2).Caption = lbl.Item(11).Caption Or lbl.Item(2).Caption = lbl.Item(18).Caption Or lbl.Item(2).Caption = lbl.Item(19).Caption Or lbl.Item(2).Caption = lbl.Item(20).Caption Or lbl.Item(2).Caption = lbl.Item(0).Caption Or lbl.Item(2).Caption = lbl.Item(1).Caption Or lbl.Item(2).Caption = lbl.Item(3).Caption Or lbl.Item(2).Caption = lbl.Item(4).Caption Or lbl.Item(2).Caption = lbl.Item(5).Caption Or lbl.Item(2).Caption = lbl.Item(6).Caption Or lbl.Item(2).Caption = lbl.Item(7).Caption Or lbl.Item(2).Caption = lbl.Item(8).Caption Or lbl.Item(2).Caption = lbl.Item(29).Caption Or lbl.Item(2).Caption = lbl.Item(32).Caption Or lbl.Item(2).Caption = lbl.Item(35).Caption Or lbl.Item(2).Caption = lbl.Item(56).Caption Or lbl.Item(2).Caption = lbl.Item(59).Caption Or lbl.Item(2).Caption = lbl.Item(62).Caption Then
    lbl.Item(2).ForeColor = &HFF&
    ElseIf lbl.Item(29).Caption = lbl.Item(2).Caption Or lbl.Item(29).Caption = lbl.Item(5).Caption Or lbl.Item(29).Caption = lbl.Item(8).Caption Or lbl.Item(29).Caption = lbl.Item(56).Caption Or lbl.Item(29).Caption = lbl.Item(59).Caption Or lbl.Item(29).Caption = lbl.Item(62).Caption Or lbl.Item(29).Caption = lbl.Item(27).Caption Or lbl.Item(29).Caption = lbl.Item(28).Caption Or lbl.Item(29).Caption = lbl.Item(30).Caption Or lbl.Item(29).Caption = lbl.Item(31).Caption Or lbl.Item(29).Caption = lbl.Item(32).Caption Or lbl.Item(29).Caption = lbl.Item(33).Caption Or lbl.Item(29).Caption = lbl.Item(34).Caption Or lbl.Item(29).Caption = lbl.Item(35).Caption Or lbl.Item(29).Caption = lbl.Item(36).Caption Or lbl.Item(29).Caption = lbl.Item(37).Caption Or lbl.Item(29).Caption = lbl.Item(38).Caption Or lbl.Item(29).Caption = lbl.Item(45).Caption Or lbl.Item(29).Caption = lbl.Item(46).Caption Or lbl.Item(29).Caption = lbl.Item(47).Caption Then
    lbl.Item(29).ForeColor = &HFF&
    ElseIf lbl.Item(30).Caption = lbl.Item(0).Caption Or lbl.Item(30).Caption = lbl.Item(3).Caption Or lbl.Item(30).Caption = lbl.Item(6).Caption Or lbl.Item(30).Caption = lbl.Item(54).Caption Or lbl.Item(30).Caption = lbl.Item(57).Caption Or lbl.Item(30).Caption = lbl.Item(60).Caption Or lbl.Item(30).Caption = lbl.Item(27).Caption Or lbl.Item(30).Caption = lbl.Item(28).Caption Or lbl.Item(30).Caption = lbl.Item(29).Caption Or lbl.Item(30).Caption = lbl.Item(31).Caption Or lbl.Item(30).Caption = lbl.Item(32).Caption Or lbl.Item(30).Caption = lbl.Item(33).Caption Or lbl.Item(30).Caption = lbl.Item(34).Caption Or lbl.Item(30).Caption = lbl.Item(35).Caption Or lbl.Item(30).Caption = lbl.Item(39).Caption Or lbl.Item(30).Caption = lbl.Item(40).Caption Or lbl.Item(30).Caption = lbl.Item(41).Caption Or lbl.Item(30).Caption = lbl.Item(48).Caption Or lbl.Item(30).Caption = lbl.Item(49).Caption Or lbl.Item(30).Caption = lbl.Item(50).Caption Then
    lbl.Item(30).ForeColor = &HFF&
    ElseIf lbl.Item(32).Caption = lbl.Item(2).Caption Or lbl.Item(32).Caption = lbl.Item(5).Caption Or lbl.Item(32).Caption = lbl.Item(8).Caption Or lbl.Item(32).Caption = lbl.Item(56).Caption Or lbl.Item(32).Caption = lbl.Item(59).Caption Or lbl.Item(32).Caption = lbl.Item(62).Caption Or lbl.Item(32).Caption = lbl.Item(27).Caption Or lbl.Item(32).Caption = lbl.Item(28).Caption Or lbl.Item(32).Caption = lbl.Item(29).Caption Or lbl.Item(32).Caption = lbl.Item(30).Caption Or lbl.Item(32).Caption = lbl.Item(31).Caption Or lbl.Item(32).Caption = lbl.Item(33).Caption Or lbl.Item(32).Caption = lbl.Item(34).Caption Or lbl.Item(32).Caption = lbl.Item(35).Caption Or lbl.Item(32).Caption = lbl.Item(39).Caption Or lbl.Item(32).Caption = lbl.Item(40).Caption Or lbl.Item(32).Caption = lbl.Item(41).Caption Or lbl.Item(32).Caption = lbl.Item(48).Caption Or lbl.Item(32).Caption = lbl.Item(49).Caption Or lbl.Item(32).Caption = lbl.Item(50).Caption Then
    lbl.Item(32).ForeColor = &HFF&
    ElseIf lbl.Item(33).Caption = lbl.Item(0).Caption Or lbl.Item(33).Caption = lbl.Item(3).Caption Or lbl.Item(33).Caption = lbl.Item(6).Caption Or lbl.Item(33).Caption = lbl.Item(54).Caption Or lbl.Item(33).Caption = lbl.Item(57).Caption Or lbl.Item(33).Caption = lbl.Item(60).Caption Or lbl.Item(33).Caption = lbl.Item(27).Caption Or lbl.Item(33).Caption = lbl.Item(28).Caption Or lbl.Item(33).Caption = lbl.Item(29).Caption Or lbl.Item(33).Caption = lbl.Item(30).Caption Or lbl.Item(33).Caption = lbl.Item(31).Caption Or lbl.Item(33).Caption = lbl.Item(32).Caption Or lbl.Item(33).Caption = lbl.Item(34).Caption Or lbl.Item(33).Caption = lbl.Item(35).Caption Or lbl.Item(33).Caption = lbl.Item(42).Caption Or lbl.Item(33).Caption = lbl.Item(43).Caption Or lbl.Item(33).Caption = lbl.Item(44).Caption Or lbl.Item(33).Caption = lbl.Item(51).Caption Or lbl.Item(33).Caption = lbl.Item(52).Caption Or lbl.Item(33).Caption = lbl.Item(53).Caption Then
    lbl.Item(33).ForeColor = &HFF&
    ElseIf lbl.Item(34).Caption = lbl.Item(1).Caption Or lbl.Item(34).Caption = lbl.Item(4).Caption Or lbl.Item(34).Caption = lbl.Item(7).Caption Or lbl.Item(34).Caption = lbl.Item(55).Caption Or lbl.Item(34).Caption = lbl.Item(58).Caption Or lbl.Item(34).Caption = lbl.Item(61).Caption Or lbl.Item(34).Caption = lbl.Item(27).Caption Or lbl.Item(34).Caption = lbl.Item(28).Caption Or lbl.Item(34).Caption = lbl.Item(29).Caption Or lbl.Item(34).Caption = lbl.Item(30).Caption Or lbl.Item(34).Caption = lbl.Item(31).Caption Or lbl.Item(34).Caption = lbl.Item(32).Caption Or lbl.Item(34).Caption = lbl.Item(33).Caption Or lbl.Item(34).Caption = lbl.Item(35).Caption Or lbl.Item(34).Caption = lbl.Item(42).Caption Or lbl.Item(34).Caption = lbl.Item(43).Caption Or lbl.Item(34).Caption = lbl.Item(44).Caption Or lbl.Item(34).Caption = lbl.Item(51).Caption Or lbl.Item(34).Caption = lbl.Item(52).Caption Or lbl.Item(34).Caption = lbl.Item(53).Caption Then
    lbl.Item(34).ForeColor = &HFF&
    ElseIf lbl.Item(35).Caption = lbl.Item(2).Caption Or lbl.Item(35).Caption = lbl.Item(5).Caption Or lbl.Item(35).Caption = lbl.Item(8).Caption Or lbl.Item(35).Caption = lbl.Item(56).Caption Or lbl.Item(35).Caption = lbl.Item(59).Caption Or lbl.Item(35).Caption = lbl.Item(62).Caption Or lbl.Item(35).Caption = lbl.Item(27).Caption Or lbl.Item(35).Caption = lbl.Item(28).Caption Or lbl.Item(35).Caption = lbl.Item(29).Caption Or lbl.Item(35).Caption = lbl.Item(30).Caption Or lbl.Item(35).Caption = lbl.Item(31).Caption Or lbl.Item(35).Caption = lbl.Item(32).Caption Or lbl.Item(35).Caption = lbl.Item(33).Caption Or lbl.Item(35).Caption = lbl.Item(34).Caption Or lbl.Item(35).Caption = lbl.Item(42).Caption Or lbl.Item(35).Caption = lbl.Item(43).Caption Or lbl.Item(35).Caption = lbl.Item(44).Caption Or lbl.Item(35).Caption = lbl.Item(51).Caption Or lbl.Item(35).Caption = lbl.Item(52).Caption Or lbl.Item(35).Caption = lbl.Item(53).Caption Then
    lbl.Item(35).ForeColor = &HFF&
    ElseIf lbl.Item(36).Caption = lbl.Item(10).Caption Or lbl.Item(36).Caption = lbl.Item(12).Caption Or lbl.Item(36).Caption = lbl.Item(15).Caption Or lbl.Item(36).Caption = lbl.Item(63).Caption Or lbl.Item(36).Caption = lbl.Item(66).Caption Or lbl.Item(36).Caption = lbl.Item(69).Caption Or lbl.Item(36).Caption = lbl.Item(37).Caption Or lbl.Item(36).Caption = lbl.Item(38).Caption Or lbl.Item(36).Caption = lbl.Item(39).Caption Or lbl.Item(36).Caption = lbl.Item(40).Caption Or lbl.Item(36).Caption = lbl.Item(41).Caption Or lbl.Item(36).Caption = lbl.Item(42).Caption Or lbl.Item(36).Caption = lbl.Item(43).Caption Or lbl.Item(36).Caption = lbl.Item(44).Caption Or lbl.Item(36).Caption = lbl.Item(27).Caption Or lbl.Item(36).Caption = lbl.Item(28).Caption Or lbl.Item(36).Caption = lbl.Item(29).Caption Or lbl.Item(36).Caption = lbl.Item(45).Caption Or lbl.Item(36).Caption = lbl.Item(46).Caption Or lbl.Item(36).Caption = lbl.Item(47).Caption Then
    lbl.Item(36).ForeColor = &HFF&
    ElseIf lbl.Item(38).Caption = lbl.Item(11).Caption Or lbl.Item(38).Caption = lbl.Item(14).Caption Or lbl.Item(38).Caption = lbl.Item(17).Caption Or lbl.Item(38).Caption = lbl.Item(65).Caption Or lbl.Item(38).Caption = lbl.Item(68).Caption Or lbl.Item(38).Caption = lbl.Item(71).Caption Or lbl.Item(38).Caption = lbl.Item(36).Caption Or lbl.Item(38).Caption = lbl.Item(37).Caption Or lbl.Item(38).Caption = lbl.Item(39).Caption Or lbl.Item(38).Caption = lbl.Item(40).Caption Or lbl.Item(38).Caption = lbl.Item(41).Caption Or lbl.Item(38).Caption = lbl.Item(42).Caption Or lbl.Item(38).Caption = lbl.Item(43).Caption Or lbl.Item(38).Caption = lbl.Item(44).Caption Or lbl.Item(38).Caption = lbl.Item(27).Caption Or lbl.Item(38).Caption = lbl.Item(28).Caption Or lbl.Item(38).Caption = lbl.Item(29).Caption Or lbl.Item(38).Caption = lbl.Item(45).Caption Or lbl.Item(38).Caption = lbl.Item(46).Caption Or lbl.Item(38).Caption = lbl.Item(47).Caption Then
    lbl.Item(38).ForeColor = &HFF&
    ElseIf lbl.Item(3).Caption = lbl.Item(12).Caption Or lbl.Item(3).Caption = lbl.Item(13).Caption Or lbl.Item(3).Caption = lbl.Item(14).Caption Or lbl.Item(3).Caption = lbl.Item(21).Caption Or lbl.Item(3).Caption = lbl.Item(22).Caption Or lbl.Item(3).Caption = lbl.Item(23).Caption Or lbl.Item(3).Caption = lbl.Item(0).Caption Or lbl.Item(3).Caption = lbl.Item(1).Caption Or lbl.Item(3).Caption = lbl.Item(2).Caption Or lbl.Item(3).Caption = lbl.Item(4).Caption Or lbl.Item(3).Caption = lbl.Item(5).Caption Or lbl.Item(3).Caption = lbl.Item(6).Caption Or lbl.Item(3).Caption = lbl.Item(7).Caption Or lbl.Item(3).Caption = lbl.Item(8).Caption Or lbl.Item(3).Caption = lbl.Item(27).Caption Or lbl.Item(3).Caption = lbl.Item(30).Caption Or lbl.Item(3).Caption = lbl.Item(33).Caption Or lbl.Item(3).Caption = lbl.Item(54).Caption Or lbl.Item(3).Caption = lbl.Item(57).Caption Or lbl.Item(3).Caption = lbl.Item(60).Caption Then
    lbl.Item(3).ForeColor = &HFF&
    ElseIf lbl.Item(39).Caption = lbl.Item(10).Caption Or lbl.Item(39).Caption = lbl.Item(12).Caption Or lbl.Item(39).Caption = lbl.Item(15).Caption Or lbl.Item(39).Caption = lbl.Item(63).Caption Or lbl.Item(39).Caption = lbl.Item(66).Caption Or lbl.Item(39).Caption = lbl.Item(69).Caption Or lbl.Item(39).Caption = lbl.Item(36).Caption Or lbl.Item(39).Caption = lbl.Item(37).Caption Or lbl.Item(39).Caption = lbl.Item(38).Caption Or lbl.Item(39).Caption = lbl.Item(40).Caption Or lbl.Item(39).Caption = lbl.Item(41).Caption Or lbl.Item(39).Caption = lbl.Item(42).Caption Or lbl.Item(39).Caption = lbl.Item(43).Caption Or lbl.Item(39).Caption = lbl.Item(44).Caption Or lbl.Item(39).Caption = lbl.Item(30).Caption Or lbl.Item(39).Caption = lbl.Item(31).Caption Or lbl.Item(39).Caption = lbl.Item(32).Caption Or lbl.Item(39).Caption = lbl.Item(48).Caption Or lbl.Item(39).Caption = lbl.Item(49).Caption Or lbl.Item(39).Caption = lbl.Item(50).Caption Then
    lbl.Item(39).ForeColor = &HFF&
    ElseIf lbl.Item(40).Caption = lbl.Item(10).Caption Or lbl.Item(40).Caption = lbl.Item(13).Caption Or lbl.Item(40).Caption = lbl.Item(16).Caption Or lbl.Item(40).Caption = lbl.Item(64).Caption Or lbl.Item(40).Caption = lbl.Item(67).Caption Or lbl.Item(40).Caption = lbl.Item(70).Caption Or lbl.Item(40).Caption = lbl.Item(36).Caption Or lbl.Item(40).Caption = lbl.Item(37).Caption Or lbl.Item(40).Caption = lbl.Item(38).Caption Or lbl.Item(40).Caption = lbl.Item(39).Caption Or lbl.Item(40).Caption = lbl.Item(41).Caption Or lbl.Item(40).Caption = lbl.Item(42).Caption Or lbl.Item(40).Caption = lbl.Item(43).Caption Or lbl.Item(40).Caption = lbl.Item(44).Caption Or lbl.Item(40).Caption = lbl.Item(30).Caption Or lbl.Item(40).Caption = lbl.Item(31).Caption Or lbl.Item(40).Caption = lbl.Item(32).Caption Or lbl.Item(40).Caption = lbl.Item(48).Caption Or lbl.Item(40).Caption = lbl.Item(49).Caption Or lbl.Item(40).Caption = lbl.Item(50).Caption Then
    lbl.Item(40).ForeColor = &HFF&
    ElseIf lbl.Item(41).Caption = lbl.Item(11).Caption Or lbl.Item(41).Caption = lbl.Item(14).Caption Or lbl.Item(41).Caption = lbl.Item(17).Caption Or lbl.Item(41).Caption = lbl.Item(65).Caption Or lbl.Item(41).Caption = lbl.Item(68).Caption Or lbl.Item(41).Caption = lbl.Item(71).Caption Or lbl.Item(41).Caption = lbl.Item(36).Caption Or lbl.Item(41).Caption = lbl.Item(37).Caption Or lbl.Item(41).Caption = lbl.Item(38).Caption Or lbl.Item(41).Caption = lbl.Item(39).Caption Or lbl.Item(41).Caption = lbl.Item(40).Caption Or lbl.Item(41).Caption = lbl.Item(42).Caption Or lbl.Item(41).Caption = lbl.Item(43).Caption Or lbl.Item(41).Caption = lbl.Item(44).Caption Or lbl.Item(41).Caption = lbl.Item(30).Caption Or lbl.Item(41).Caption = lbl.Item(31).Caption Or lbl.Item(41).Caption = lbl.Item(32).Caption Or lbl.Item(41).Caption = lbl.Item(48).Caption Or lbl.Item(41).Caption = lbl.Item(49).Caption Or lbl.Item(41).Caption = lbl.Item(50).Caption Then
    lbl.Item(41).ForeColor = &HFF&
    ElseIf lbl.Item(42).Caption = lbl.Item(10).Caption Or lbl.Item(42).Caption = lbl.Item(12).Caption Or lbl.Item(42).Caption = lbl.Item(15).Caption Or lbl.Item(42).Caption = lbl.Item(63).Caption Or lbl.Item(42).Caption = lbl.Item(66).Caption Or lbl.Item(42).Caption = lbl.Item(69).Caption Or lbl.Item(42).Caption = lbl.Item(36).Caption Or lbl.Item(42).Caption = lbl.Item(37).Caption Or lbl.Item(42).Caption = lbl.Item(38).Caption Or lbl.Item(42).Caption = lbl.Item(39).Caption Or lbl.Item(42).Caption = lbl.Item(40).Caption Or lbl.Item(42).Caption = lbl.Item(41).Caption Or lbl.Item(42).Caption = lbl.Item(43).Caption Or lbl.Item(42).Caption = lbl.Item(44).Caption Or lbl.Item(42).Caption = lbl.Item(33).Caption Or lbl.Item(42).Caption = lbl.Item(34).Caption Or lbl.Item(42).Caption = lbl.Item(35).Caption Or lbl.Item(42).Caption = lbl.Item(51).Caption Or lbl.Item(42).Caption = lbl.Item(52).Caption Or lbl.Item(42).Caption = lbl.Item(53).Caption Then
    lbl.Item(42).ForeColor = &HFF&
    ElseIf lbl.Item(43).Caption = lbl.Item(10).Caption Or lbl.Item(43).Caption = lbl.Item(13).Caption Or lbl.Item(43).Caption = lbl.Item(16).Caption Or lbl.Item(43).Caption = lbl.Item(64).Caption Or lbl.Item(43).Caption = lbl.Item(67).Caption Or lbl.Item(43).Caption = lbl.Item(70).Caption Or lbl.Item(43).Caption = lbl.Item(36).Caption Or lbl.Item(43).Caption = lbl.Item(37).Caption Or lbl.Item(43).Caption = lbl.Item(38).Caption Or lbl.Item(43).Caption = lbl.Item(39).Caption Or lbl.Item(43).Caption = lbl.Item(40).Caption Or lbl.Item(43).Caption = lbl.Item(41).Caption Or lbl.Item(43).Caption = lbl.Item(42).Caption Or lbl.Item(43).Caption = lbl.Item(44).Caption Or lbl.Item(43).Caption = lbl.Item(33).Caption Or lbl.Item(43).Caption = lbl.Item(34).Caption Or lbl.Item(43).Caption = lbl.Item(35).Caption Or lbl.Item(43).Caption = lbl.Item(51).Caption Or lbl.Item(43).Caption = lbl.Item(52).Caption Or lbl.Item(43).Caption = lbl.Item(53).Caption Then
    lbl.Item(43).ForeColor = &HFF&
    ElseIf lbl.Item(44).Caption = lbl.Item(11).Caption Or lbl.Item(44).Caption = lbl.Item(14).Caption Or lbl.Item(44).Caption = lbl.Item(17).Caption Or lbl.Item(44).Caption = lbl.Item(65).Caption Or lbl.Item(44).Caption = lbl.Item(68).Caption Or lbl.Item(44).Caption = lbl.Item(71).Caption Or lbl.Item(44).Caption = lbl.Item(36).Caption Or lbl.Item(44).Caption = lbl.Item(37).Caption Or lbl.Item(44).Caption = lbl.Item(38).Caption Or lbl.Item(44).Caption = lbl.Item(39).Caption Or lbl.Item(44).Caption = lbl.Item(40).Caption Or lbl.Item(44).Caption = lbl.Item(41).Caption Or lbl.Item(44).Caption = lbl.Item(42).Caption Or lbl.Item(44).Caption = lbl.Item(43).Caption Or lbl.Item(44).Caption = lbl.Item(33).Caption Or lbl.Item(44).Caption = lbl.Item(34).Caption Or lbl.Item(44).Caption = lbl.Item(35).Caption Or lbl.Item(44).Caption = lbl.Item(51).Caption Or lbl.Item(44).Caption = lbl.Item(52).Caption Or lbl.Item(44).Caption = lbl.Item(53).Caption Then
    lbl.Item(44).ForeColor = &HFF&
    ElseIf lbl.Item(45).Caption = lbl.Item(18).Caption Or lbl.Item(45).Caption = lbl.Item(21).Caption Or lbl.Item(45).Caption = lbl.Item(24).Caption Or lbl.Item(45).Caption = lbl.Item(72).Caption Or lbl.Item(45).Caption = lbl.Item(75).Caption Or lbl.Item(45).Caption = lbl.Item(78).Caption Or lbl.Item(45).Caption = lbl.Item(46).Caption Or lbl.Item(45).Caption = lbl.Item(47).Caption Or lbl.Item(45).Caption = lbl.Item(48).Caption Or lbl.Item(45).Caption = lbl.Item(49).Caption Or lbl.Item(45).Caption = lbl.Item(50).Caption Or lbl.Item(45).Caption = lbl.Item(51).Caption Or lbl.Item(45).Caption = lbl.Item(52).Caption Or lbl.Item(45).Caption = lbl.Item(53).Caption Or lbl.Item(45).Caption = lbl.Item(27).Caption Or lbl.Item(45).Caption = lbl.Item(28).Caption Or lbl.Item(45).Caption = lbl.Item(29).Caption Or lbl.Item(45).Caption = lbl.Item(36).Caption Or lbl.Item(45).Caption = lbl.Item(37).Caption Or lbl.Item(45).Caption = lbl.Item(38).Caption Then
    lbl.Item(45).ForeColor = &HFF&
    ElseIf lbl.Item(46).Caption = lbl.Item(19).Caption Or lbl.Item(46).Caption = lbl.Item(22).Caption Or lbl.Item(46).Caption = lbl.Item(25).Caption Or lbl.Item(46).Caption = lbl.Item(73).Caption Or lbl.Item(46).Caption = lbl.Item(76).Caption Or lbl.Item(46).Caption = lbl.Item(79).Caption Or lbl.Item(46).Caption = lbl.Item(45).Caption Or lbl.Item(46).Caption = lbl.Item(47).Caption Or lbl.Item(46).Caption = lbl.Item(48).Caption Or lbl.Item(46).Caption = lbl.Item(49).Caption Or lbl.Item(46).Caption = lbl.Item(50).Caption Or lbl.Item(46).Caption = lbl.Item(51).Caption Or lbl.Item(46).Caption = lbl.Item(52).Caption Or lbl.Item(46).Caption = lbl.Item(53).Caption Or lbl.Item(46).Caption = lbl.Item(27).Caption Or lbl.Item(46).Caption = lbl.Item(28).Caption Or lbl.Item(46).Caption = lbl.Item(29).Caption Or lbl.Item(46).Caption = lbl.Item(36).Caption Or lbl.Item(46).Caption = lbl.Item(37).Caption Or lbl.Item(46).Caption = lbl.Item(38).Caption Then
    lbl.Item(46).ForeColor = &HFF&
    ElseIf lbl.Item(47).Caption = lbl.Item(20).Caption Or lbl.Item(47).Caption = lbl.Item(23).Caption Or lbl.Item(47).Caption = lbl.Item(26).Caption Or lbl.Item(47).Caption = lbl.Item(74).Caption Or lbl.Item(47).Caption = lbl.Item(77).Caption Or lbl.Item(47).Caption = lbl.Item(80).Caption Or lbl.Item(47).Caption = lbl.Item(45).Caption Or lbl.Item(47).Caption = lbl.Item(46).Caption Or lbl.Item(47).Caption = lbl.Item(48).Caption Or lbl.Item(47).Caption = lbl.Item(49).Caption Or lbl.Item(47).Caption = lbl.Item(50).Caption Or lbl.Item(47).Caption = lbl.Item(51).Caption Or lbl.Item(47).Caption = lbl.Item(52).Caption Or lbl.Item(47).Caption = lbl.Item(53).Caption Or lbl.Item(47).Caption = lbl.Item(27).Caption Or lbl.Item(47).Caption = lbl.Item(28).Caption Or lbl.Item(47).Caption = lbl.Item(29).Caption Or lbl.Item(47).Caption = lbl.Item(36).Caption Or lbl.Item(47).Caption = lbl.Item(37).Caption Or lbl.Item(47).Caption = lbl.Item(38).Caption Then
    lbl.Item(47).ForeColor = &HFF&
    ElseIf lbl.Item(48).Caption = lbl.Item(18).Caption Or lbl.Item(48).Caption = lbl.Item(21).Caption Or lbl.Item(48).Caption = lbl.Item(24).Caption Or lbl.Item(48).Caption = lbl.Item(72).Caption Or lbl.Item(48).Caption = lbl.Item(75).Caption Or lbl.Item(48).Caption = lbl.Item(78).Caption Or lbl.Item(48).Caption = lbl.Item(45).Caption Or lbl.Item(48).Caption = lbl.Item(46).Caption Or lbl.Item(48).Caption = lbl.Item(47).Caption Or lbl.Item(48).Caption = lbl.Item(49).Caption Or lbl.Item(48).Caption = lbl.Item(50).Caption Or lbl.Item(48).Caption = lbl.Item(51).Caption Or lbl.Item(48).Caption = lbl.Item(52).Caption Or lbl.Item(48).Caption = lbl.Item(53).Caption Or lbl.Item(48).Caption = lbl.Item(30).Caption Or lbl.Item(48).Caption = lbl.Item(31).Caption Or lbl.Item(48).Caption = lbl.Item(32).Caption Or lbl.Item(48).Caption = lbl.Item(39).Caption Or lbl.Item(48).Caption = lbl.Item(40).Caption Or lbl.Item(48).Caption = lbl.Item(41).Caption Then
    lbl.Item(48).ForeColor = &HFF&
    ElseIf lbl.Item(4).Caption = lbl.Item(28).Caption Or lbl.Item(4).Caption = lbl.Item(31).Caption Or lbl.Item(4).Caption = lbl.Item(34).Caption Or lbl.Item(4).Caption = lbl.Item(55).Caption Or lbl.Item(4).Caption = lbl.Item(58).Caption Or lbl.Item(4).Caption = lbl.Item(61).Caption Or lbl.Item(4).Caption = lbl.Item(0).Caption Or lbl.Item(4).Caption = lbl.Item(1).Caption Or lbl.Item(4).Caption = lbl.Item(2).Caption Or lbl.Item(4).Caption = lbl.Item(3).Caption Or lbl.Item(4).Caption = lbl.Item(5).Caption Or lbl.Item(4).Caption = lbl.Item(6).Caption Or lbl.Item(4).Caption = lbl.Item(7).Caption Or lbl.Item(4).Caption = lbl.Item(8).Caption Or lbl.Item(4).Caption = lbl.Item(12).Caption Or lbl.Item(4).Caption = lbl.Item(13).Caption Or lbl.Item(4).Caption = lbl.Item(0).Caption Or lbl.Item(4).Caption = lbl.Item(21).Caption Or lbl.Item(4).Caption = lbl.Item(22).Caption Or lbl.Item(4).Caption = lbl.Item(23).Caption Then
    lbl.Item(4).ForeColor = &HFF&
    ElseIf lbl.Item(49).Caption = lbl.Item(19).Caption Or lbl.Item(49).Caption = lbl.Item(22).Caption Or lbl.Item(49).Caption = lbl.Item(25).Caption Or lbl.Item(49).Caption = lbl.Item(73).Caption Or lbl.Item(49).Caption = lbl.Item(76).Caption Or lbl.Item(49).Caption = lbl.Item(79).Caption Or lbl.Item(49).Caption = lbl.Item(45).Caption Or lbl.Item(49).Caption = lbl.Item(46).Caption Or lbl.Item(49).Caption = lbl.Item(47).Caption Or lbl.Item(49).Caption = lbl.Item(48).Caption Or lbl.Item(49).Caption = lbl.Item(50).Caption Or lbl.Item(49).Caption = lbl.Item(51).Caption Or lbl.Item(49).Caption = lbl.Item(52).Caption Or lbl.Item(49).Caption = lbl.Item(53).Caption Or lbl.Item(49).Caption = lbl.Item(30).Caption Or lbl.Item(49).Caption = lbl.Item(31).Caption Or lbl.Item(49).Caption = lbl.Item(32).Caption Or lbl.Item(49).Caption = lbl.Item(39).Caption Or lbl.Item(49).Caption = lbl.Item(40).Caption Or lbl.Item(49).Caption = lbl.Item(41).Caption Then
    lbl.Item(49).ForeColor = &HFF&
    ElseIf lbl.Item(50).Caption = lbl.Item(20).Caption Or lbl.Item(50).Caption = lbl.Item(23).Caption Or lbl.Item(50).Caption = lbl.Item(26).Caption Or lbl.Item(50).Caption = lbl.Item(74).Caption Or lbl.Item(50).Caption = lbl.Item(77).Caption Or lbl.Item(50).Caption = lbl.Item(80).Caption Or lbl.Item(50).Caption = lbl.Item(45).Caption Or lbl.Item(50).Caption = lbl.Item(46).Caption Or lbl.Item(50).Caption = lbl.Item(47).Caption Or lbl.Item(50).Caption = lbl.Item(48).Caption Or lbl.Item(50).Caption = lbl.Item(49).Caption Or lbl.Item(50).Caption = lbl.Item(51).Caption Or lbl.Item(50).Caption = lbl.Item(52).Caption Or lbl.Item(50).Caption = lbl.Item(53).Caption Or lbl.Item(50).Caption = lbl.Item(30).Caption Or lbl.Item(50).Caption = lbl.Item(31).Caption Or lbl.Item(50).Caption = lbl.Item(32).Caption Or lbl.Item(50).Caption = lbl.Item(39).Caption Or lbl.Item(50).Caption = lbl.Item(40).Caption Or lbl.Item(50).Caption = lbl.Item(41).Caption Then
    lbl.Item(50).ForeColor = &HFF&
    ElseIf lbl.Item(52).Caption = lbl.Item(19).Caption Or lbl.Item(52).Caption = lbl.Item(22).Caption Or lbl.Item(52).Caption = lbl.Item(25).Caption Or lbl.Item(52).Caption = lbl.Item(73).Caption Or lbl.Item(52).Caption = lbl.Item(76).Caption Or lbl.Item(52).Caption = lbl.Item(79).Caption Or lbl.Item(52).Caption = lbl.Item(45).Caption Or lbl.Item(52).Caption = lbl.Item(46).Caption Or lbl.Item(52).Caption = lbl.Item(47).Caption Or lbl.Item(52).Caption = lbl.Item(48).Caption Or lbl.Item(52).Caption = lbl.Item(49).Caption Or lbl.Item(52).Caption = lbl.Item(50).Caption Or lbl.Item(52).Caption = lbl.Item(51).Caption Or lbl.Item(52).Caption = lbl.Item(53).Caption Or lbl.Item(52).Caption = lbl.Item(33).Caption Or lbl.Item(52).Caption = lbl.Item(34).Caption Or lbl.Item(52).Caption = lbl.Item(35).Caption Or lbl.Item(52).Caption = lbl.Item(42).Caption Or lbl.Item(52).Caption = lbl.Item(43).Caption Or lbl.Item(52).Caption = lbl.Item(44).Caption Then
    lbl.Item(52).ForeColor = &HFF&
    ElseIf lbl.Item(53).Caption = lbl.Item(20).Caption Or lbl.Item(53).Caption = lbl.Item(23).Caption Or lbl.Item(53).Caption = lbl.Item(26).Caption Or lbl.Item(53).Caption = lbl.Item(74).Caption Or lbl.Item(53).Caption = lbl.Item(77).Caption Or lbl.Item(53).Caption = lbl.Item(80).Caption Or lbl.Item(53).Caption = lbl.Item(45).Caption Or lbl.Item(53).Caption = lbl.Item(46).Caption Or lbl.Item(53).Caption = lbl.Item(47).Caption Or lbl.Item(53).Caption = lbl.Item(48).Caption Or lbl.Item(53).Caption = lbl.Item(49).Caption Or lbl.Item(53).Caption = lbl.Item(50).Caption Or lbl.Item(53).Caption = lbl.Item(51).Caption Or lbl.Item(53).Caption = lbl.Item(52).Caption Or lbl.Item(53).Caption = lbl.Item(33).Caption Or lbl.Item(53).Caption = lbl.Item(34).Caption Or lbl.Item(53).Caption = lbl.Item(35).Caption Or lbl.Item(53).Caption = lbl.Item(42).Caption Or lbl.Item(53).Caption = lbl.Item(43).Caption Or lbl.Item(53).Caption = lbl.Item(44).Caption Then
    lbl.Item(53).ForeColor = &HFF&
    ElseIf lbl.Item(54).Caption = lbl.Item(0).Caption Or lbl.Item(54).Caption = lbl.Item(3).Caption Or lbl.Item(54).Caption = lbl.Item(6).Caption Or lbl.Item(54).Caption = lbl.Item(27).Caption Or lbl.Item(54).Caption = lbl.Item(30).Caption Or lbl.Item(54).Caption = lbl.Item(33).Caption Or lbl.Item(54).Caption = lbl.Item(55).Caption Or lbl.Item(54).Caption = lbl.Item(56).Caption Or lbl.Item(54).Caption = lbl.Item(57).Caption Or lbl.Item(54).Caption = lbl.Item(58).Caption Or lbl.Item(54).Caption = lbl.Item(59).Caption Or lbl.Item(54).Caption = lbl.Item(60).Caption Or lbl.Item(54).Caption = lbl.Item(61).Caption Or lbl.Item(54).Caption = lbl.Item(62).Caption Or lbl.Item(54).Caption = lbl.Item(63).Caption Or lbl.Item(54).Caption = lbl.Item(64).Caption Or lbl.Item(54).Caption = lbl.Item(65).Caption Or lbl.Item(54).Caption = lbl.Item(72).Caption Or lbl.Item(54).Caption = lbl.Item(73).Caption Or lbl.Item(54).Caption = lbl.Item(74).Caption Then
    lbl.Item(54).ForeColor = &HFF&
    ElseIf lbl.Item(55).Caption = lbl.Item(1).Caption Or lbl.Item(55).Caption = lbl.Item(4).Caption Or lbl.Item(55).Caption = lbl.Item(7).Caption Or lbl.Item(55).Caption = lbl.Item(28).Caption Or lbl.Item(55).Caption = lbl.Item(31).Caption Or lbl.Item(55).Caption = lbl.Item(34).Caption Or lbl.Item(55).Caption = lbl.Item(54).Caption Or lbl.Item(55).Caption = lbl.Item(56).Caption Or lbl.Item(55).Caption = lbl.Item(57).Caption Or lbl.Item(55).Caption = lbl.Item(58).Caption Or lbl.Item(55).Caption = lbl.Item(59).Caption Or lbl.Item(55).Caption = lbl.Item(60).Caption Or lbl.Item(55).Caption = lbl.Item(61).Caption Or lbl.Item(55).Caption = lbl.Item(62).Caption Or lbl.Item(55).Caption = lbl.Item(63).Caption Or lbl.Item(55).Caption = lbl.Item(64).Caption Or lbl.Item(55).Caption = lbl.Item(65).Caption Or lbl.Item(55).Caption = lbl.Item(72).Caption Or lbl.Item(55).Caption = lbl.Item(73).Caption Or lbl.Item(55).Caption = lbl.Item(74).Caption Then
    lbl.Item(55).ForeColor = &HFF&
    ElseIf lbl.Item(56).Caption = lbl.Item(2).Caption Or lbl.Item(56).Caption = lbl.Item(5).Caption Or lbl.Item(56).Caption = lbl.Item(8).Caption Or lbl.Item(56).Caption = lbl.Item(29).Caption Or lbl.Item(56).Caption = lbl.Item(32).Caption Or lbl.Item(56).Caption = lbl.Item(35).Caption Or lbl.Item(56).Caption = lbl.Item(54).Caption Or lbl.Item(56).Caption = lbl.Item(55).Caption Or lbl.Item(56).Caption = lbl.Item(57).Caption Or lbl.Item(56).Caption = lbl.Item(58).Caption Or lbl.Item(56).Caption = lbl.Item(59).Caption Or lbl.Item(56).Caption = lbl.Item(60).Caption Or lbl.Item(56).Caption = lbl.Item(61).Caption Or lbl.Item(56).Caption = lbl.Item(62).Caption Or lbl.Item(56).Caption = lbl.Item(63).Caption Or lbl.Item(56).Caption = lbl.Item(64).Caption Or lbl.Item(56).Caption = lbl.Item(65).Caption Or lbl.Item(56).Caption = lbl.Item(72).Caption Or lbl.Item(56).Caption = lbl.Item(73).Caption Or lbl.Item(56).Caption = lbl.Item(74).Caption Then
    lbl.Item(56).ForeColor = &HFF&
    ElseIf lbl.Item(57).Caption = lbl.Item(0).Caption Or lbl.Item(57).Caption = lbl.Item(3).Caption Or lbl.Item(57).Caption = lbl.Item(6).Caption Or lbl.Item(57).Caption = lbl.Item(27).Caption Or lbl.Item(57).Caption = lbl.Item(30).Caption Or lbl.Item(57).Caption = lbl.Item(33).Caption Or lbl.Item(57).Caption = lbl.Item(54).Caption Or lbl.Item(57).Caption = lbl.Item(55).Caption Or lbl.Item(57).Caption = lbl.Item(56).Caption Or lbl.Item(57).Caption = lbl.Item(58).Caption Or lbl.Item(57).Caption = lbl.Item(59).Caption Or lbl.Item(57).Caption = lbl.Item(60).Caption Or lbl.Item(57).Caption = lbl.Item(61).Caption Or lbl.Item(57).Caption = lbl.Item(62).Caption Or lbl.Item(57).Caption = lbl.Item(66).Caption Or lbl.Item(57).Caption = lbl.Item(67).Caption Or lbl.Item(57).Caption = lbl.Item(68).Caption Or lbl.Item(57).Caption = lbl.Item(75).Caption Or lbl.Item(57).Caption = lbl.Item(76).Caption Or lbl.Item(57).Caption = lbl.Item(77).Caption Then
    lbl.Item(57).ForeColor = &HFF&
    ElseIf lbl.Item(58).Caption = lbl.Item(1).Caption Or lbl.Item(58).Caption = lbl.Item(4).Caption Or lbl.Item(58).Caption = lbl.Item(7).Caption Or lbl.Item(58).Caption = lbl.Item(28).Caption Or lbl.Item(58).Caption = lbl.Item(31).Caption Or lbl.Item(58).Caption = lbl.Item(34).Caption Or lbl.Item(58).Caption = lbl.Item(54).Caption Or lbl.Item(58).Caption = lbl.Item(55).Caption Or lbl.Item(58).Caption = lbl.Item(56).Caption Or lbl.Item(58).Caption = lbl.Item(57).Caption Or lbl.Item(58).Caption = lbl.Item(59).Caption Or lbl.Item(58).Caption = lbl.Item(60).Caption Or lbl.Item(58).Caption = lbl.Item(61).Caption Or lbl.Item(58).Caption = lbl.Item(62).Caption Or lbl.Item(58).Caption = lbl.Item(66).Caption Or lbl.Item(58).Caption = lbl.Item(67).Caption Or lbl.Item(58).Caption = lbl.Item(68).Caption Or lbl.Item(58).Caption = lbl.Item(75).Caption Or lbl.Item(58).Caption = lbl.Item(76).Caption Or lbl.Item(58).Caption = lbl.Item(77).Caption Then
    lbl.Item(58).ForeColor = &HFF&
    ElseIf lbl.Item(5).Caption = lbl.Item(29).Caption Or lbl.Item(5).Caption = lbl.Item(32).Caption Or lbl.Item(5).Caption = lbl.Item(35).Caption Or lbl.Item(5).Caption = lbl.Item(56).Caption Or lbl.Item(5).Caption = lbl.Item(59).Caption Or lbl.Item(5).Caption = lbl.Item(62).Caption Or lbl.Item(5).Caption = lbl.Item(0).Caption Or lbl.Item(5).Caption = lbl.Item(1).Caption Or lbl.Item(5).Caption = lbl.Item(2).Caption Or lbl.Item(5).Caption = lbl.Item(3).Caption Or lbl.Item(5).Caption = lbl.Item(4).Caption Or lbl.Item(5).Caption = lbl.Item(6).Caption Or lbl.Item(5).Caption = lbl.Item(7).Caption Or lbl.Item(5).Caption = lbl.Item(8).Caption Or lbl.Item(5).Caption = lbl.Item(12).Caption Or lbl.Item(5).Caption = lbl.Item(13).Caption Or lbl.Item(5).Caption = lbl.Item(14).Caption Or lbl.Item(5).Caption = lbl.Item(21).Caption Or lbl.Item(5).Caption = lbl.Item(22).Caption Or lbl.Item(5).Caption = lbl.Item(23).Caption Then
    lbl.Item(5).ForeColor = &HFF&
    ElseIf lbl.Item(59).Caption = lbl.Item(2).Caption Or lbl.Item(59).Caption = lbl.Item(5).Caption Or lbl.Item(59).Caption = lbl.Item(8).Caption Or lbl.Item(59).Caption = lbl.Item(29).Caption Or lbl.Item(59).Caption = lbl.Item(32).Caption Or lbl.Item(59).Caption = lbl.Item(35).Caption Or lbl.Item(59).Caption = lbl.Item(54).Caption Or lbl.Item(59).Caption = lbl.Item(55).Caption Or lbl.Item(59).Caption = lbl.Item(56).Caption Or lbl.Item(59).Caption = lbl.Item(57).Caption Or lbl.Item(59).Caption = lbl.Item(58).Caption Or lbl.Item(59).Caption = lbl.Item(60).Caption Or lbl.Item(59).Caption = lbl.Item(61).Caption Or lbl.Item(59).Caption = lbl.Item(62).Caption Or lbl.Item(59).Caption = lbl.Item(66).Caption Or lbl.Item(59).Caption = lbl.Item(67).Caption Or lbl.Item(59).Caption = lbl.Item(68).Caption Or lbl.Item(59).Caption = lbl.Item(75).Caption Or lbl.Item(59).Caption = lbl.Item(76).Caption Or lbl.Item(59).Caption = lbl.Item(77).Caption Then
    lbl.Item(59).ForeColor = &HFF&
    ElseIf lbl.Item(61).Caption = lbl.Item(1).Caption Or lbl.Item(61).Caption = lbl.Item(4).Caption Or lbl.Item(61).Caption = lbl.Item(7).Caption Or lbl.Item(61).Caption = lbl.Item(28).Caption Or lbl.Item(61).Caption = lbl.Item(31).Caption Or lbl.Item(61).Caption = lbl.Item(34).Caption Or lbl.Item(61).Caption = lbl.Item(54).Caption Or lbl.Item(61).Caption = lbl.Item(55).Caption Or lbl.Item(61).Caption = lbl.Item(56).Caption Or lbl.Item(61).Caption = lbl.Item(57).Caption Or lbl.Item(61).Caption = lbl.Item(58).Caption Or lbl.Item(61).Caption = lbl.Item(59).Caption Or lbl.Item(61).Caption = lbl.Item(60).Caption Or lbl.Item(61).Caption = lbl.Item(62).Caption Or lbl.Item(61).Caption = lbl.Item(69).Caption Or lbl.Item(61).Caption = lbl.Item(70).Caption Or lbl.Item(61).Caption = lbl.Item(71).Caption Or lbl.Item(61).Caption = lbl.Item(78).Caption Or lbl.Item(61).Caption = lbl.Item(79).Caption Or lbl.Item(61).Caption = lbl.Item(80).Caption Then
    lbl.Item(61).ForeColor = &HFF&
    ElseIf lbl.Item(62).Caption = lbl.Item(2).Caption Or lbl.Item(62).Caption = lbl.Item(5).Caption Or lbl.Item(62).Caption = lbl.Item(8).Caption Or lbl.Item(62).Caption = lbl.Item(29).Caption Or lbl.Item(62).Caption = lbl.Item(32).Caption Or lbl.Item(62).Caption = lbl.Item(35).Caption Or lbl.Item(62).Caption = lbl.Item(54).Caption Or lbl.Item(62).Caption = lbl.Item(55).Caption Or lbl.Item(62).Caption = lbl.Item(56).Caption Or lbl.Item(62).Caption = lbl.Item(57).Caption Or lbl.Item(62).Caption = lbl.Item(58).Caption Or lbl.Item(62).Caption = lbl.Item(59).Caption Or lbl.Item(62).Caption = lbl.Item(60).Caption Or lbl.Item(62).Caption = lbl.Item(61).Caption Or lbl.Item(62).Caption = lbl.Item(69).Caption Or lbl.Item(62).Caption = lbl.Item(70).Caption Or lbl.Item(62).Caption = lbl.Item(71).Caption Or lbl.Item(62).Caption = lbl.Item(78).Caption Or lbl.Item(62).Caption = lbl.Item(79).Caption Or lbl.Item(62).Caption = lbl.Item(80).Caption Then
    lbl.Item(62).ForeColor = &HFF&
    ElseIf lbl.Item(63).Caption = lbl.Item(10).Caption Or lbl.Item(63).Caption = lbl.Item(12).Caption Or lbl.Item(63).Caption = lbl.Item(15).Caption Or lbl.Item(63).Caption = lbl.Item(36).Caption Or lbl.Item(63).Caption = lbl.Item(39).Caption Or lbl.Item(63).Caption = lbl.Item(42).Caption Or lbl.Item(63).Caption = lbl.Item(64).Caption Or lbl.Item(63).Caption = lbl.Item(65).Caption Or lbl.Item(63).Caption = lbl.Item(66).Caption Or lbl.Item(63).Caption = lbl.Item(67).Caption Or lbl.Item(63).Caption = lbl.Item(68).Caption Or lbl.Item(63).Caption = lbl.Item(69).Caption Or lbl.Item(63).Caption = lbl.Item(70).Caption Or lbl.Item(63).Caption = lbl.Item(71).Caption Or lbl.Item(63).Caption = lbl.Item(54).Caption Or lbl.Item(63).Caption = lbl.Item(55).Caption Or lbl.Item(63).Caption = lbl.Item(56).Caption Or lbl.Item(63).Caption = lbl.Item(72).Caption Or lbl.Item(63).Caption = lbl.Item(73).Caption Or lbl.Item(63).Caption = lbl.Item(74).Caption Then
    lbl.Item(63).ForeColor = &HFF&
    ElseIf lbl.Item(64).Caption = lbl.Item(10).Caption Or lbl.Item(64).Caption = lbl.Item(13).Caption Or lbl.Item(64).Caption = lbl.Item(16).Caption Or lbl.Item(64).Caption = lbl.Item(37).Caption Or lbl.Item(64).Caption = lbl.Item(40).Caption Or lbl.Item(64).Caption = lbl.Item(43).Caption Or lbl.Item(64).Caption = lbl.Item(63).Caption Or lbl.Item(64).Caption = lbl.Item(65).Caption Or lbl.Item(64).Caption = lbl.Item(66).Caption Or lbl.Item(64).Caption = lbl.Item(67).Caption Or lbl.Item(64).Caption = lbl.Item(68).Caption Or lbl.Item(64).Caption = lbl.Item(69).Caption Or lbl.Item(64).Caption = lbl.Item(70).Caption Or lbl.Item(64).Caption = lbl.Item(71).Caption Or lbl.Item(64).Caption = lbl.Item(54).Caption Or lbl.Item(64).Caption = lbl.Item(55).Caption Or lbl.Item(64).Caption = lbl.Item(56).Caption Or lbl.Item(64).Caption = lbl.Item(72).Caption Or lbl.Item(64).Caption = lbl.Item(73).Caption Or lbl.Item(64).Caption = lbl.Item(74).Caption Then
    lbl.Item(64).ForeColor = &HFF&
    ElseIf lbl.Item(65).Caption = lbl.Item(11).Caption Or lbl.Item(65).Caption = lbl.Item(14).Caption Or lbl.Item(65).Caption = lbl.Item(17).Caption Or lbl.Item(65).Caption = lbl.Item(38).Caption Or lbl.Item(65).Caption = lbl.Item(41).Caption Or lbl.Item(65).Caption = lbl.Item(44).Caption Or lbl.Item(65).Caption = lbl.Item(63).Caption Or lbl.Item(65).Caption = lbl.Item(64).Caption Or lbl.Item(65).Caption = lbl.Item(66).Caption Or lbl.Item(65).Caption = lbl.Item(67).Caption Or lbl.Item(65).Caption = lbl.Item(68).Caption Or lbl.Item(65).Caption = lbl.Item(69).Caption Or lbl.Item(65).Caption = lbl.Item(70).Caption Or lbl.Item(65).Caption = lbl.Item(71).Caption Or lbl.Item(65).Caption = lbl.Item(54).Caption Or lbl.Item(65).Caption = lbl.Item(55).Caption Or lbl.Item(65).Caption = lbl.Item(56).Caption Or lbl.Item(65).Caption = lbl.Item(72).Caption Or lbl.Item(65).Caption = lbl.Item(73).Caption Or lbl.Item(65).Caption = lbl.Item(74).Caption Then
    lbl.Item(65).ForeColor = &HFF&
    ElseIf lbl.Item(66).Caption = lbl.Item(10).Caption Or lbl.Item(66).Caption = lbl.Item(12).Caption Or lbl.Item(66).Caption = lbl.Item(15).Caption Or lbl.Item(66).Caption = lbl.Item(36).Caption Or lbl.Item(66).Caption = lbl.Item(39).Caption Or lbl.Item(66).Caption = lbl.Item(42).Caption Or lbl.Item(66).Caption = lbl.Item(63).Caption Or lbl.Item(66).Caption = lbl.Item(64).Caption Or lbl.Item(66).Caption = lbl.Item(65).Caption Or lbl.Item(66).Caption = lbl.Item(67).Caption Or lbl.Item(66).Caption = lbl.Item(68).Caption Or lbl.Item(66).Caption = lbl.Item(69).Caption Or lbl.Item(66).Caption = lbl.Item(70).Caption Or lbl.Item(66).Caption = lbl.Item(71).Caption Or lbl.Item(66).Caption = lbl.Item(57).Caption Or lbl.Item(66).Caption = lbl.Item(58).Caption Or lbl.Item(66).Caption = lbl.Item(59).Caption Or lbl.Item(66).Caption = lbl.Item(75).Caption Or lbl.Item(66).Caption = lbl.Item(76).Caption Or lbl.Item(66).Caption = lbl.Item(77).Caption Then
    lbl.Item(66).ForeColor = &HFF&
    ElseIf lbl.Item(68).Caption = lbl.Item(11).Caption Or lbl.Item(68).Caption = lbl.Item(14).Caption Or lbl.Item(68).Caption = lbl.Item(17).Caption Or lbl.Item(68).Caption = lbl.Item(38).Caption Or lbl.Item(68).Caption = lbl.Item(41).Caption Or lbl.Item(68).Caption = lbl.Item(44).Caption Or lbl.Item(68).Caption = lbl.Item(63).Caption Or lbl.Item(68).Caption = lbl.Item(64).Caption Or lbl.Item(68).Caption = lbl.Item(65).Caption Or lbl.Item(68).Caption = lbl.Item(66).Caption Or lbl.Item(68).Caption = lbl.Item(67).Caption Or lbl.Item(68).Caption = lbl.Item(69).Caption Or lbl.Item(68).Caption = lbl.Item(70).Caption Or lbl.Item(68).Caption = lbl.Item(71).Caption Or lbl.Item(68).Caption = lbl.Item(57).Caption Or lbl.Item(68).Caption = lbl.Item(58).Caption Or lbl.Item(68).Caption = lbl.Item(59).Caption Or lbl.Item(68).Caption = lbl.Item(75).Caption Or lbl.Item(68).Caption = lbl.Item(76).Caption Or lbl.Item(68).Caption = lbl.Item(77).Caption Then
    lbl.Item(68).ForeColor = &HFF&
    ElseIf lbl.Item(6).Caption = lbl.Item(27).Caption Or lbl.Item(6).Caption = lbl.Item(30).Caption Or lbl.Item(6).Caption = lbl.Item(33).Caption Or lbl.Item(6).Caption = lbl.Item(54).Caption Or lbl.Item(6).Caption = lbl.Item(57).Caption Or lbl.Item(6).Caption = lbl.Item(60).Caption Or lbl.Item(6).Caption = lbl.Item(0).Caption Or lbl.Item(6).Caption = lbl.Item(1).Caption Or lbl.Item(6).Caption = lbl.Item(2).Caption Or lbl.Item(6).Caption = lbl.Item(3).Caption Or lbl.Item(6).Caption = lbl.Item(4).Caption Or lbl.Item(6).Caption = lbl.Item(5).Caption Or lbl.Item(6).Caption = lbl.Item(7).Caption Or lbl.Item(6).Caption = lbl.Item(8).Caption Or lbl.Item(6).Caption = lbl.Item(15).Caption Or lbl.Item(6).Caption = lbl.Item(16).Caption Or lbl.Item(6).Caption = lbl.Item(17).Caption Or lbl.Item(6).Caption = lbl.Item(24).Caption Or lbl.Item(6).Caption = lbl.Item(25).Caption Or lbl.Item(6).Caption = lbl.Item(26).Caption Then
    lbl.Item(6).ForeColor = &HFF&
    ElseIf lbl.Item(69).Caption = lbl.Item(10).Caption Or lbl.Item(69).Caption = lbl.Item(12).Caption Or lbl.Item(69).Caption = lbl.Item(15).Caption Or lbl.Item(69).Caption = lbl.Item(36).Caption Or lbl.Item(69).Caption = lbl.Item(39).Caption Or lbl.Item(69).Caption = lbl.Item(42).Caption Or lbl.Item(69).Caption = lbl.Item(63).Caption Or lbl.Item(69).Caption = lbl.Item(64).Caption Or lbl.Item(69).Caption = lbl.Item(65).Caption Or lbl.Item(69).Caption = lbl.Item(66).Caption Or lbl.Item(69).Caption = lbl.Item(67).Caption Or lbl.Item(69).Caption = lbl.Item(68).Caption Or lbl.Item(69).Caption = lbl.Item(70).Caption Or lbl.Item(69).Caption = lbl.Item(71).Caption Or lbl.Item(69).Caption = lbl.Item(60).Caption Or lbl.Item(69).Caption = lbl.Item(61).Caption Or lbl.Item(69).Caption = lbl.Item(62).Caption Or lbl.Item(69).Caption = lbl.Item(78).Caption Or lbl.Item(69).Caption = lbl.Item(79).Caption Or lbl.Item(69).Caption = lbl.Item(80).Caption Then
    lbl.Item(69).ForeColor = &HFF&
    ElseIf lbl.Item(70).Caption = lbl.Item(10).Caption Or lbl.Item(70).Caption = lbl.Item(13).Caption Or lbl.Item(70).Caption = lbl.Item(16).Caption Or lbl.Item(70).Caption = lbl.Item(37).Caption Or lbl.Item(70).Caption = lbl.Item(40).Caption Or lbl.Item(70).Caption = lbl.Item(43).Caption Or lbl.Item(70).Caption = lbl.Item(63).Caption Or lbl.Item(70).Caption = lbl.Item(64).Caption Or lbl.Item(70).Caption = lbl.Item(65).Caption Or lbl.Item(70).Caption = lbl.Item(66).Caption Or lbl.Item(70).Caption = lbl.Item(67).Caption Or lbl.Item(70).Caption = lbl.Item(68).Caption Or lbl.Item(70).Caption = lbl.Item(69).Caption Or lbl.Item(70).Caption = lbl.Item(71).Caption Or lbl.Item(70).Caption = lbl.Item(60).Caption Or lbl.Item(70).Caption = lbl.Item(61).Caption Or lbl.Item(70).Caption = lbl.Item(62).Caption Or lbl.Item(70).Caption = lbl.Item(78).Caption Or lbl.Item(70).Caption = lbl.Item(79).Caption Or lbl.Item(70).Caption = lbl.Item(80).Caption Then
    lbl.Item(70).ForeColor = &HFF&
    ElseIf lbl.Item(71).Caption = lbl.Item(11).Caption Or lbl.Item(71).Caption = lbl.Item(14).Caption Or lbl.Item(71).Caption = lbl.Item(17).Caption Or lbl.Item(71).Caption = lbl.Item(38).Caption Or lbl.Item(71).Caption = lbl.Item(41).Caption Or lbl.Item(71).Caption = lbl.Item(44).Caption Or lbl.Item(71).Caption = lbl.Item(63).Caption Or lbl.Item(71).Caption = lbl.Item(64).Caption Or lbl.Item(71).Caption = lbl.Item(65).Caption Or lbl.Item(71).Caption = lbl.Item(66).Caption Or lbl.Item(71).Caption = lbl.Item(67).Caption Or lbl.Item(71).Caption = lbl.Item(68).Caption Or lbl.Item(71).Caption = lbl.Item(69).Caption Or lbl.Item(71).Caption = lbl.Item(70).Caption Or lbl.Item(71).Caption = lbl.Item(60).Caption Or lbl.Item(71).Caption = lbl.Item(61).Caption Or lbl.Item(71).Caption = lbl.Item(62).Caption Or lbl.Item(71).Caption = lbl.Item(78).Caption Or lbl.Item(71).Caption = lbl.Item(79).Caption Or lbl.Item(71).Caption = lbl.Item(80).Caption Then
    lbl.Item(71).ForeColor = &HFF&
    ElseIf lbl.Item(72).Caption = lbl.Item(18).Caption Or lbl.Item(72).Caption = lbl.Item(21).Caption Or lbl.Item(72).Caption = lbl.Item(24).Caption Or lbl.Item(72).Caption = lbl.Item(45).Caption Or lbl.Item(72).Caption = lbl.Item(48).Caption Or lbl.Item(72).Caption = lbl.Item(51).Caption Or lbl.Item(72).Caption = lbl.Item(73).Caption Or lbl.Item(72).Caption = lbl.Item(74).Caption Or lbl.Item(72).Caption = lbl.Item(75).Caption Or lbl.Item(72).Caption = lbl.Item(76).Caption Or lbl.Item(72).Caption = lbl.Item(77).Caption Or lbl.Item(72).Caption = lbl.Item(78).Caption Or lbl.Item(72).Caption = lbl.Item(79).Caption Or lbl.Item(72).Caption = lbl.Item(80).Caption Or lbl.Item(72).Caption = lbl.Item(54).Caption Or lbl.Item(72).Caption = lbl.Item(55).Caption Or lbl.Item(72).Caption = lbl.Item(56).Caption Or lbl.Item(72).Caption = lbl.Item(63).Caption Or lbl.Item(72).Caption = lbl.Item(64).Caption Or lbl.Item(72).Caption = lbl.Item(65).Caption Then
    lbl.Item(72).ForeColor = &HFF&
    ElseIf lbl.Item(73).Caption = lbl.Item(19).Caption Or lbl.Item(73).Caption = lbl.Item(22).Caption Or lbl.Item(73).Caption = lbl.Item(25).Caption Or lbl.Item(73).Caption = lbl.Item(46).Caption Or lbl.Item(73).Caption = lbl.Item(49).Caption Or lbl.Item(73).Caption = lbl.Item(52).Caption Or lbl.Item(73).Caption = lbl.Item(72).Caption Or lbl.Item(73).Caption = lbl.Item(74).Caption Or lbl.Item(73).Caption = lbl.Item(75).Caption Or lbl.Item(73).Caption = lbl.Item(76).Caption Or lbl.Item(73).Caption = lbl.Item(77).Caption Or lbl.Item(73).Caption = lbl.Item(78).Caption Or lbl.Item(73).Caption = lbl.Item(79).Caption Or lbl.Item(73).Caption = lbl.Item(80).Caption Or lbl.Item(73).Caption = lbl.Item(54).Caption Or lbl.Item(73).Caption = lbl.Item(55).Caption Or lbl.Item(73).Caption = lbl.Item(56).Caption Or lbl.Item(73).Caption = lbl.Item(63).Caption Or lbl.Item(73).Caption = lbl.Item(64).Caption Or lbl.Item(73).Caption = lbl.Item(65).Caption Then
    lbl.Item(73).ForeColor = &HFF&
    ElseIf lbl.Item(75).Caption = lbl.Item(18).Caption Or lbl.Item(75).Caption = lbl.Item(21).Caption Or lbl.Item(75).Caption = lbl.Item(24).Caption Or lbl.Item(75).Caption = lbl.Item(45).Caption Or lbl.Item(75).Caption = lbl.Item(48).Caption Or lbl.Item(75).Caption = lbl.Item(51).Caption Or lbl.Item(75).Caption = lbl.Item(72).Caption Or lbl.Item(75).Caption = lbl.Item(73).Caption Or lbl.Item(75).Caption = lbl.Item(74).Caption Or lbl.Item(75).Caption = lbl.Item(76).Caption Or lbl.Item(75).Caption = lbl.Item(77).Caption Or lbl.Item(75).Caption = lbl.Item(78).Caption Or lbl.Item(75).Caption = lbl.Item(79).Caption Or lbl.Item(75).Caption = lbl.Item(80).Caption Or lbl.Item(75).Caption = lbl.Item(57).Caption Or lbl.Item(75).Caption = lbl.Item(58).Caption Or lbl.Item(75).Caption = lbl.Item(59).Caption Or lbl.Item(75).Caption = lbl.Item(66).Caption Or lbl.Item(75).Caption = lbl.Item(67).Caption Or lbl.Item(75).Caption = lbl.Item(68).Caption Then
    lbl.Item(75).ForeColor = &HFF&
    ElseIf lbl.Item(76).Caption = lbl.Item(19).Caption Or lbl.Item(76).Caption = lbl.Item(22).Caption Or lbl.Item(76).Caption = lbl.Item(25).Caption Or lbl.Item(76).Caption = lbl.Item(46).Caption Or lbl.Item(76).Caption = lbl.Item(49).Caption Or lbl.Item(76).Caption = lbl.Item(52).Caption Or lbl.Item(76).Caption = lbl.Item(72).Caption Or lbl.Item(76).Caption = lbl.Item(73).Caption Or lbl.Item(76).Caption = lbl.Item(74).Caption Or lbl.Item(76).Caption = lbl.Item(75).Caption Or lbl.Item(76).Caption = lbl.Item(77).Caption Or lbl.Item(76).Caption = lbl.Item(78).Caption Or lbl.Item(76).Caption = lbl.Item(79).Caption Or lbl.Item(76).Caption = lbl.Item(80).Caption Or lbl.Item(76).Caption = lbl.Item(57).Caption Or lbl.Item(76).Caption = lbl.Item(58).Caption Or lbl.Item(76).Caption = lbl.Item(59).Caption Or lbl.Item(76).Caption = lbl.Item(66).Caption Or lbl.Item(76).Caption = lbl.Item(67).Caption Or lbl.Item(76).Caption = lbl.Item(68).Caption Then
    lbl.Item(76).ForeColor = &HFF&
    ElseIf lbl.Item(77).Caption = lbl.Item(20).Caption Or lbl.Item(77).Caption = lbl.Item(23).Caption Or lbl.Item(77).Caption = lbl.Item(26).Caption Or lbl.Item(77).Caption = lbl.Item(47).Caption Or lbl.Item(77).Caption = lbl.Item(50).Caption Or lbl.Item(77).Caption = lbl.Item(53).Caption Or lbl.Item(77).Caption = lbl.Item(72).Caption Or lbl.Item(77).Caption = lbl.Item(73).Caption Or lbl.Item(77).Caption = lbl.Item(74).Caption Or lbl.Item(77).Caption = lbl.Item(75).Caption Or lbl.Item(77).Caption = lbl.Item(76).Caption Or lbl.Item(77).Caption = lbl.Item(78).Caption Or lbl.Item(77).Caption = lbl.Item(79).Caption Or lbl.Item(77).Caption = lbl.Item(80).Caption Or lbl.Item(77).Caption = lbl.Item(57).Caption Or lbl.Item(77).Caption = lbl.Item(58).Caption Or lbl.Item(77).Caption = lbl.Item(59).Caption Or lbl.Item(77).Caption = lbl.Item(66).Caption Or lbl.Item(77).Caption = lbl.Item(67).Caption Or lbl.Item(77).Caption = lbl.Item(68).Caption Then
    lbl.Item(77).ForeColor = &HFF&
    ElseIf lbl.Item(78).Caption = lbl.Item(18).Caption Or lbl.Item(78).Caption = lbl.Item(21).Caption Or lbl.Item(78).Caption = lbl.Item(24).Caption Or lbl.Item(78).Caption = lbl.Item(45).Caption Or lbl.Item(78).Caption = lbl.Item(48).Caption Or lbl.Item(78).Caption = lbl.Item(51).Caption Or lbl.Item(78).Caption = lbl.Item(72).Caption Or lbl.Item(78).Caption = lbl.Item(73).Caption Or lbl.Item(78).Caption = lbl.Item(74).Caption Or lbl.Item(78).Caption = lbl.Item(75).Caption Or lbl.Item(78).Caption = lbl.Item(76).Caption Or lbl.Item(78).Caption = lbl.Item(77).Caption Or lbl.Item(78).Caption = lbl.Item(79).Caption Or lbl.Item(78).Caption = lbl.Item(80).Caption Or lbl.Item(78).Caption = lbl.Item(60).Caption Or lbl.Item(78).Caption = lbl.Item(61).Caption Or lbl.Item(78).Caption = lbl.Item(62).Caption Or lbl.Item(78).Caption = lbl.Item(69).Caption Or lbl.Item(78).Caption = lbl.Item(70).Caption Or lbl.Item(78).Caption = lbl.Item(71).Caption Then
    lbl.Item(78).ForeColor = &HFF&
    ElseIf lbl.Item(7).Caption = lbl.Item(28).Caption Or lbl.Item(7).Caption = lbl.Item(31).Caption Or lbl.Item(7).Caption = lbl.Item(34).Caption Or lbl.Item(7).Caption = lbl.Item(55).Caption Or lbl.Item(7).Caption = lbl.Item(58).Caption Or lbl.Item(7).Caption = lbl.Item(61).Caption Or lbl.Item(7).Caption = lbl.Item(0).Caption Or lbl.Item(7).Caption = lbl.Item(1).Caption Or lbl.Item(7).Caption = lbl.Item(2).Caption Or lbl.Item(7).Caption = lbl.Item(3).Caption Or lbl.Item(7).Caption = lbl.Item(4).Caption Or lbl.Item(7).Caption = lbl.Item(5).Caption Or lbl.Item(7).Caption = lbl.Item(6).Caption Or lbl.Item(7).Caption = lbl.Item(8).Caption Or lbl.Item(7).Caption = lbl.Item(15).Caption Or lbl.Item(7).Caption = lbl.Item(16).Caption Or lbl.Item(7).Caption = lbl.Item(17).Caption Or lbl.Item(7).Caption = lbl.Item(24).Caption Or lbl.Item(7).Caption = lbl.Item(25).Caption Or lbl.Item(7).Caption = lbl.Item(26).Caption Then
    lbl.Item(7).ForeColor = &HFF&
    ElseIf lbl.Item(79).Caption = lbl.Item(19).Caption Or lbl.Item(79).Caption = lbl.Item(22).Caption Or lbl.Item(79).Caption = lbl.Item(25).Caption Or lbl.Item(79).Caption = lbl.Item(46).Caption Or lbl.Item(79).Caption = lbl.Item(49).Caption Or lbl.Item(79).Caption = lbl.Item(52).Caption Or lbl.Item(79).Caption = lbl.Item(72).Caption Or lbl.Item(79).Caption = lbl.Item(73).Caption Or lbl.Item(79).Caption = lbl.Item(74).Caption Or lbl.Item(79).Caption = lbl.Item(75).Caption Or lbl.Item(79).Caption = lbl.Item(76).Caption Or lbl.Item(79).Caption = lbl.Item(77).Caption Or lbl.Item(79).Caption = lbl.Item(78).Caption Or lbl.Item(79).Caption = lbl.Item(80).Caption Or lbl.Item(79).Caption = lbl.Item(60).Caption Or lbl.Item(79).Caption = lbl.Item(61).Caption Or lbl.Item(79).Caption = lbl.Item(62).Caption Or lbl.Item(79).Caption = lbl.Item(69).Caption Or lbl.Item(79).Caption = lbl.Item(70).Caption Or lbl.Item(79).Caption = lbl.Item(71).Caption Then
    lbl.Item(79).ForeColor = &HFF&
    ElseIf lbl.Item(80).Caption = lbl.Item(20).Caption Or lbl.Item(80).Caption = lbl.Item(23).Caption Or lbl.Item(80).Caption = lbl.Item(26).Caption Or lbl.Item(80).Caption = lbl.Item(47).Caption Or lbl.Item(80).Caption = lbl.Item(50).Caption Or lbl.Item(80).Caption = lbl.Item(53).Caption Or lbl.Item(80).Caption = lbl.Item(72).Caption Or lbl.Item(80).Caption = lbl.Item(73).Caption Or lbl.Item(80).Caption = lbl.Item(74).Caption Or lbl.Item(80).Caption = lbl.Item(75).Caption Or lbl.Item(80).Caption = lbl.Item(76).Caption Or lbl.Item(80).Caption = lbl.Item(77).Caption Or lbl.Item(80).Caption = lbl.Item(78).Caption Or lbl.Item(80).Caption = lbl.Item(79).Caption Or lbl.Item(80).Caption = lbl.Item(60).Caption Or lbl.Item(80).Caption = lbl.Item(61).Caption Or lbl.Item(80).Caption = lbl.Item(62).Caption Or lbl.Item(80).Caption = lbl.Item(69).Caption Or lbl.Item(80).Caption = lbl.Item(70).Caption Or lbl.Item(80).Caption = lbl.Item(71).Caption Then
    lbl.Item(80).ForeColor = &HFF&
    ElseIf lbl.Item(8).Caption = lbl.Item(29).Caption Or lbl.Item(8).Caption = lbl.Item(32).Caption Or lbl.Item(8).Caption = lbl.Item(35).Caption Or lbl.Item(8).Caption = lbl.Item(56).Caption Or lbl.Item(8).Caption = lbl.Item(59).Caption Or lbl.Item(8).Caption = lbl.Item(62).Caption Or lbl.Item(8).Caption = lbl.Item(0).Caption Or lbl.Item(8).Caption = lbl.Item(1).Caption Or lbl.Item(8).Caption = lbl.Item(2).Caption Or lbl.Item(8).Caption = lbl.Item(3).Caption Or lbl.Item(8).Caption = lbl.Item(4).Caption Or lbl.Item(8).Caption = lbl.Item(5).Caption Or lbl.Item(8).Caption = lbl.Item(6).Caption Or lbl.Item(8).Caption = lbl.Item(7).Caption Or lbl.Item(8).Caption = lbl.Item(15).Caption Or lbl.Item(8).Caption = lbl.Item(16).Caption Or lbl.Item(8).Caption = lbl.Item(17).Caption Or lbl.Item(8).Caption = lbl.Item(24).Caption Or lbl.Item(8).Caption = lbl.Item(25).Caption Or lbl.Item(8).Caption = lbl.Item(26).Caption Then
    lbl.Item(8).ForeColor = &HFF&
    End If
End Sub
Sub Limpar()
For A = 0 To 80
lbl.Item(A).Caption = ""
Next
End Sub
