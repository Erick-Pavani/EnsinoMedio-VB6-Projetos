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
    Else
    lbl.Item(0).ForeColor = &H0&
	lbl.Item(10).Caption = cmbNumeros.Text
ElseIf lbl.Item(10).Caption = lbl.Item(10).CaptionOr lbl.Item(10).Caption = lbl.Item(11).CaptionOr lbl.Item(10).Caption = lbl.Item(0).CaptionOr lbl.Item(10).Caption = lbl.Item(1).CaptionOr lbl.Item(10).Caption = lbl.Item(2).CaptionOr lbl.Item(10).Caption = lbl.Item(18).CaptionOr lbl.Item(10).Caption = lbl.Item(19).CaptionOr lbl.Item(10).Caption = lbl.Item(20).CaptionOr lbl.Item(10).Caption = lbl.Item(12).CaptionOr lbl.Item(10).Caption = lbl.Item(15).CaptionOr lbl.Item(10).Caption = lbl.Item(36).Caption Or lbl.Item(10).Caption = lbl.Item(39).Caption Or lbl.Item(10).Caption = lbl.Item(42).Caption Or lbl.Item(10).Caption = lbl.Item(63).Caption Or lbl.Item(10).Caption = lbl.Item(66).Caption Or lbl.Item(10).Caption = lbl.Item(69).Caption Or lbl.Item(10).Caption = lbl.Item(13).CaptionOr lbl.Item(10).Caption = lbl.Item(14).CaptionOr lbl.Item(10).Caption = lbl.Item(16).CaptionOr lbl.Item(10).Caption = lbl.Item(17).CaptionThen
    lbl10.ForeColor = &HFF&
Else
    lbl10.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl11_Click()
lbl.Item(10).Caption= cmbNumeros.Text
If lbl.Item(10).Caption= lbl.Item(0).CaptionOr lbl.Item(10).Caption= lbl.Item(1).CaptionOr lbl.Item(10).Caption= lbl.Item(2).CaptionOr lbl.Item(10).Caption= lbl.Item(10).Caption Or lbl.Item(10).Caption= lbl.Item(11).CaptionOr lbl.Item(10).Caption= lbl.Item(18).CaptionOr lbl.Item(10).Caption= lbl.Item(19).CaptionOr lbl.Item(10).Caption= lbl.Item(20).CaptionOr lbl.Item(10).Caption= lbl.Item(13).CaptionOr lbl.Item(10).Caption= lbl.Item(16).CaptionOr lbl.Item(10).Caption= lbl.Item(37).Caption Or lbl.Item(10).Caption= lbl.Item(40).Caption Or lbl.Item(10).Caption= lbl.Item(43).Caption Or lbl.Item(10).Caption= lbl.Item(64).Caption Or lbl.Item(10).Caption= lbl.Item(67).Caption Or lbl.Item(10).Caption= lbl.Item(70).Caption Or lbl.Item(10).Caption= lbl.Item(12).CaptionOr lbl.Item(10).Caption= lbl.Item(14).CaptionOr lbl.Item(10).Caption= lbl.Item(15).CaptionOr lbl.Item(10).Caption= lbl.Item(17).CaptionThen
    lbl11.ForeColor = &HFF&
Else
    lbl11.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl12_Click()
lbl.Item(11).Caption= cmbNumeros.Text
If lbl.Item(11).Caption= lbl.Item(0).CaptionOr lbl.Item(11).Caption= lbl.Item(1).CaptionOr lbl.Item(11).Caption= lbl.Item(2).CaptionOr lbl.Item(11).Caption= lbl.Item(10).Caption Or lbl.Item(11).Caption= lbl.Item(10).CaptionOr lbl.Item(11).Caption= lbl.Item(18).CaptionOr lbl.Item(11).Caption= lbl.Item(19).CaptionOr lbl.Item(11).Caption= lbl.Item(20).CaptionOr lbl.Item(11).Caption= lbl.Item(14).CaptionOr lbl.Item(11).Caption= lbl.Item(17).CaptionOr lbl.Item(11).Caption= lbl.Item(38).Caption Or lbl.Item(11).Caption= lbl.Item(41).Caption Or lbl.Item(11).Caption= lbl.Item(44).Caption Or lbl.Item(11).Caption= lbl.Item(65).Caption Or lbl.Item(11).Caption= lbl.Item(68).Caption Or lbl.Item(11).Caption= lbl.Item(71).Caption Or lbl.Item(11).Caption= lbl.Item(12).CaptionOr lbl.Item(11).Caption= lbl.Item(13).CaptionOr lbl.Item(11).Caption= lbl.Item(15).CaptionOr lbl.Item(11).Caption= lbl.Item(16).CaptionThen
    lbl12.ForeColor = &HFF&
Else
    lbl12.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl13_Click()
lbl.Item(12).Caption= cmbNumeros.Text
If lbl.Item(12).Caption= lbl.Item(3).CaptionOr lbl.Item(12).Caption= lbl.Item(4).CaptionOr lbl.Item(12).Caption= lbl.Item(5).CaptionOr lbl.Item(12).Caption= lbl.Item(13).CaptionOr lbl.Item(12).Caption= lbl.Item(14).CaptionOr lbl.Item(12).Caption= lbl.Item(21).CaptionOr lbl.Item(12).Caption= lbl.Item(22).CaptionOr lbl.Item(12).Caption= lbl.Item(23).CaptionOr lbl.Item(12).Caption= lbl.Item(10).Caption Or lbl.Item(12).Caption= lbl.Item(15).CaptionOr lbl.Item(12).Caption= lbl.Item(36).Caption Or lbl.Item(12).Caption= lbl.Item(39).Caption Or lbl.Item(12).Caption= lbl.Item(42).Caption Or lbl.Item(12).Caption= lbl.Item(63).Caption Or lbl.Item(12).Caption= lbl.Item(66).Caption Or lbl.Item(12).Caption= lbl.Item(69).Caption Or lbl.Item(12).Caption= lbl.Item(10).CaptionOr lbl.Item(12).Caption= lbl.Item(11).CaptionOr lbl.Item(12).Caption= lbl.Item(16).CaptionOr lbl.Item(12).Caption= lbl.Item(17).CaptionThen
    lbl13.ForeColor = &HFF&
Else
    lbl13.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl15_Click()
lbl.Item(14).Caption= cmbNumeros.Text
If lbl.Item(14).Caption= lbl.Item(10).Caption Or lbl.Item(14).Caption= lbl.Item(10).CaptionOr lbl.Item(14).Caption= lbl.Item(11).CaptionOr lbl.Item(14).Caption= lbl.Item(12).CaptionOr lbl.Item(14).Caption= lbl.Item(13).CaptionOr lbl.Item(14).Caption= lbl.Item(15).CaptionOr lbl.Item(14).Caption= lbl.Item(16).CaptionOr lbl.Item(14).Caption= lbl.Item(17).CaptionOr lbl.Item(14).Caption= lbl.Item(3).CaptionOr lbl.Item(14).Caption= lbl.Item(4).CaptionOr lbl.Item(14).Caption= lbl.Item(5).CaptionOr lbl.Item(14).Caption= lbl.Item(21).CaptionOr lbl.Item(14).Caption= lbl.Item(22).CaptionOr lbl.Item(14).Caption= lbl.Item(23).CaptionOr lbl.Item(14).Caption= lbl.Item(38).Caption Or lbl.Item(14).Caption= lbl.Item(41).Caption Or lbl.Item(14).Caption= lbl.Item(44).Caption Or lbl.Item(14).Caption= lbl.Item(65).Caption Or lbl.Item(14).Caption= lbl.Item(68).Caption Or lbl.Item(14).Caption= lbl.Item(71).Caption Then
    lbl15.ForeColor = &HFF&
Else
    lbl15.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl16_Click()
lbl.Item(15).Caption= cmbNumeros.Text
If lbl.Item(15).Caption= lbl.Item(10).Caption Or lbl.Item(15).Caption= lbl.Item(10).CaptionOr lbl.Item(15).Caption= lbl.Item(11).CaptionOr lbl.Item(15).Caption= lbl.Item(12).CaptionOr lbl.Item(15).Caption= lbl.Item(13).CaptionOr lbl.Item(15).Caption= lbl.Item(14).CaptionOr lbl.Item(15).Caption= lbl.Item(16).CaptionOr lbl.Item(15).Caption= lbl.Item(17).CaptionOr lbl.Item(15).Caption= lbl.Item(6).CaptionOr lbl.Item(15).Caption= lbl.Item(7).CaptionOr lbl.Item(15).Caption= lbl.Item(8).CaptionOr lbl.Item(15).Caption= lbl.Item(24).CaptionOr lbl.Item(15).Caption= lbl.Item(25).Caption Or lbl.Item(15).Caption= lbl.Item(26).Caption Or lbl.Item(15).Caption= lbl.Item(36).Caption Or lbl.Item(15).Caption= lbl.Item(39).Caption Or lbl.Item(15).Caption= lbl.Item(42).Caption Or lbl.Item(15).Caption= lbl.Item(63).Caption Or lbl.Item(15).Caption= lbl.Item(66).Caption Or lbl.Item(15).Caption= lbl.Item(69).Caption Then
    lbl16.ForeColor = &HFF&
Else
    lbl16.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl17_Click()
lbl.Item(16).Caption= cmbNumeros.Text
If lbl.Item(16).Caption= lbl.Item(6).CaptionOr lbl.Item(16).Caption= lbl.Item(7).CaptionOr lbl.Item(16).Caption= lbl.Item(8).CaptionOr lbl.Item(16).Caption= lbl.Item(24).CaptionOr lbl.Item(16).Caption= lbl.Item(25).Caption Or lbl.Item(16).Caption= lbl.Item(26).Caption Or lbl.Item(16).Caption= lbl.Item(10).Caption Or lbl.Item(16).Caption= lbl.Item(10).CaptionOr lbl.Item(16).Caption= lbl.Item(11).CaptionOr lbl.Item(16).Caption= lbl.Item(12).CaptionOr lbl.Item(16).Caption= lbl.Item(13).CaptionOr lbl.Item(16).Caption= lbl.Item(14).CaptionOr lbl.Item(16).Caption= lbl.Item(15).CaptionOr lbl.Item(16).Caption= lbl.Item(17).CaptionOr lbl.Item(16).Caption= lbl.Item(37).Caption Or lbl.Item(16).Caption= lbl.Item(40).Caption Or lbl.Item(16).Caption= lbl.Item(43).Caption Or lbl.Item(16).Caption= lbl.Item(64).Caption Or lbl.Item(16).Caption= lbl.Item(67).Caption Or lbl.Item(16).Caption= lbl.Item(70).Caption Then
    lbl17.ForeColor = &HFF&
Else
    lbl17.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl18_Click()
lbl.Item(17).Caption= cmbNumeros.Text
If lbl.Item(17).Caption= lbl.Item(6).CaptionOr lbl.Item(17).Caption= lbl.Item(7).CaptionOr lbl.Item(17).Caption= lbl.Item(8).CaptionOr lbl.Item(17).Caption= lbl.Item(24).CaptionOr lbl.Item(17).Caption= lbl.Item(25).Caption Or lbl.Item(17).Caption= lbl.Item(26).Caption Or lbl.Item(17).Caption= lbl.Item(10).Caption Or lbl.Item(17).Caption= lbl.Item(10).CaptionOr lbl.Item(17).Caption= lbl.Item(11).CaptionOr lbl.Item(17).Caption= lbl.Item(12).CaptionOr lbl.Item(17).Caption= lbl.Item(13).CaptionOr lbl.Item(17).Caption= lbl.Item(14).CaptionOr lbl.Item(17).Caption= lbl.Item(15).CaptionOr lbl.Item(17).Caption= lbl.Item(16).CaptionOr lbl.Item(17).Caption= lbl.Item(38).Caption Or lbl.Item(17).Caption= lbl.Item(41).Caption Or lbl.Item(17).Caption= lbl.Item(44).Caption Or lbl.Item(17).Caption= lbl.Item(65).Caption Or lbl.Item(17).Caption= lbl.Item(68).Caption Or lbl.Item(17).Caption= lbl.Item(71).Caption Then
    lbl18.ForeColor = &HFF&
Else
    lbl18.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl19_Click()
lbl.Item(18).Caption= cmbNumeros.Text
If lbl.Item(18).Caption= lbl.Item(0).CaptionOr lbl.Item(18).Caption= lbl.Item(1).CaptionOr lbl.Item(18).Caption= lbl.Item(2).CaptionOr lbl.Item(18).Caption= lbl.Item(10).Caption Or lbl.Item(18).Caption= lbl.Item(10).CaptionOr lbl.Item(18).Caption= lbl.Item(11).CaptionOr lbl.Item(18).Caption= lbl.Item(19).CaptionOr lbl.Item(18).Caption= lbl.Item(20).CaptionOr lbl.Item(18).Caption= lbl.Item(21).CaptionOr lbl.Item(18).Caption= lbl.Item(22).CaptionOr lbl.Item(18).Caption= lbl.Item(23).CaptionOr lbl.Item(18).Caption= lbl.Item(24).CaptionOr lbl.Item(18).Caption= lbl.Item(25).Caption Or lbl.Item(18).Caption= lbl.Item(26).Caption Or lbl.Item(18).Caption= lbl.Item(45).Caption Or lbl.Item(18).Caption= lbl.Item(48).Caption Or lbl.Item(18).Caption= lbl.Item(51).Caption Or lbl.Item(18).Caption= lbl.Item(72).Caption Or lbl.Item(18).Caption= lbl.Item(75).Caption Or lbl.Item(18).Caption= lbl.Item(78).Caption Then
    lbl19.ForeColor = &HFF&
Else
    lbl18.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl2_Click()
lbl.Item(1).Caption= cmbNumeros.Text
If lbl.Item(1).Caption= lbl.Item(10).Caption Or lbl.Item(1).Caption= lbl.Item(10).CaptionOr lbl.Item(1).Caption= lbl.Item(11).CaptionOr lbl.Item(1).Caption= lbl.Item(18).CaptionOr lbl.Item(1).Caption= lbl.Item(19).CaptionOr lbl.Item(1).Caption= lbl.Item(20).CaptionOr lbl.Item(1).Caption= lbl.Item(0).CaptionOr lbl.Item(1).Caption= lbl.Item(2).CaptionOr lbl.Item(1).Caption= lbl.Item(3).CaptionOr lbl.Item(1).Caption= lbl.Item(4).CaptionOr lbl.Item(1).Caption= lbl.Item(5).CaptionOr lbl.Item(1).Caption= lbl.Item(6).CaptionOr lbl.Item(1).Caption= lbl.Item(7).CaptionOr lbl.Item(1).Caption= lbl.Item(8).CaptionOr lbl.Item(1).Caption= lbl.Item(28).Caption Or lbl.Item(1).Caption= lbl.Item(31).Caption Or lbl.Item(1).Caption= lbl.Item(34).Caption Or lbl.Item(1).Caption= lbl.Item(55).Caption Or lbl.Item(1).Caption= lbl.Item(58).Caption Or lbl.Item(1).Caption= lbl.Item(61).Caption Then
    lbl2.ForeColor = &HFF&
Else
    lbl2.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl20_Click()
lbl.Item(19).Caption= cmbNumeros.Text
If lbl.Item(19).Caption= lbl.Item(0).CaptionOr lbl.Item(19).Caption= lbl.Item(1).CaptionOr lbl.Item(19).Caption= lbl.Item(2).CaptionOr lbl.Item(19).Caption= lbl.Item(10).Caption Or lbl.Item(19).Caption= lbl.Item(10).CaptionOr lbl.Item(19).Caption= lbl.Item(11).CaptionOr lbl.Item(19).Caption= lbl.Item(18).CaptionOr lbl.Item(19).Caption= lbl.Item(20).CaptionOr lbl.Item(19).Caption= lbl.Item(21).CaptionOr lbl.Item(19).Caption= lbl.Item(22).CaptionOr lbl.Item(19).Caption= lbl.Item(23).CaptionOr lbl.Item(19).Caption= lbl.Item(24).CaptionOr lbl.Item(19).Caption= lbl.Item(25).Caption Or lbl.Item(19).Caption= lbl.Item(26).Caption Or lbl.Item(19).Caption= lbl.Item(46).Caption Or lbl.Item(19).Caption= lbl.Item(49).Caption Or lbl.Item(19).Caption= lbl.Item(52).Caption Or lbl.Item(19).Caption= lbl.Item(73).Caption Or lbl.Item(19).Caption= lbl.Item(76).Caption Or lbl.Item(19).Caption= lbl.Item(79).Caption Then
    lbl20.ForeColor = &HFF&
Else
    lbl20.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl21_Click()
lbl.Item(20).Caption= cmbNumeros.Text
If lbl.Item(20).Caption= lbl.Item(0).CaptionOr lbl.Item(20).Caption= lbl.Item(1).CaptionOr lbl.Item(20).Caption= lbl.Item(2).CaptionOr lbl.Item(20).Caption= lbl.Item(10).Caption Or lbl.Item(20).Caption= lbl.Item(10).CaptionOr lbl.Item(20).Caption= lbl.Item(11).CaptionOr lbl.Item(20).Caption= lbl.Item(18).CaptionOr lbl.Item(20).Caption= lbl.Item(19).CaptionOr lbl.Item(20).Caption= lbl.Item(21).CaptionOr lbl.Item(20).Caption= lbl.Item(22).CaptionOr lbl.Item(20).Caption= lbl.Item(23).CaptionOr lbl.Item(20).Caption= lbl.Item(24).CaptionOr lbl.Item(20).Caption= lbl.Item(25).Caption Or lbl.Item(20).Caption= lbl.Item(26).Caption Or lbl.Item(20).Caption= lbl.Item(47).Caption Or lbl.Item(20).Caption= lbl.Item(50).Caption Or lbl.Item(20).Caption= lbl.Item(53).Caption Or lbl.Item(20).Caption= lbl.Item(74).Caption Or lbl.Item(20).Caption= lbl.Item(77).Caption Or lbl.Item(20).Caption= lbl.Item(80).Caption Then
    lbl21.ForeColor = &HFF&
Else
    lbl21.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl22_Click()
lbl.Item(21).Caption= cmbNumeros.Text
If lbl.Item(21).Caption= lbl.Item(3).CaptionOr lbl.Item(21).Caption= lbl.Item(4).CaptionOr lbl.Item(21).Caption= lbl.Item(5).CaptionOr lbl.Item(21).Caption= lbl.Item(12).CaptionOr lbl.Item(21).Caption= lbl.Item(13).CaptionOr lbl.Item(21).Caption= lbl.Item(14).CaptionOr lbl.Item(21).Caption= lbl.Item(18).CaptionOr lbl.Item(21).Caption= lbl.Item(19).CaptionOr lbl.Item(21).Caption= lbl.Item(20).CaptionOr lbl.Item(21).Caption= lbl.Item(22).CaptionOr lbl.Item(21).Caption= lbl.Item(23).CaptionOr lbl.Item(21).Caption= lbl.Item(24).CaptionOr lbl.Item(21).Caption= lbl.Item(25).Caption Or lbl.Item(21).Caption= lbl.Item(26).Caption Or lbl.Item(21).Caption= lbl.Item(45).Caption Or lbl.Item(21).Caption= lbl.Item(48).Caption Or lbl.Item(21).Caption= lbl.Item(51).Caption Or lbl.Item(21).Caption= lbl.Item(72).Caption Or lbl.Item(21).Caption= lbl.Item(75).Caption Or lbl.Item(21).Caption= lbl.Item(78).Caption Then
    lbl22.ForeColor = &HFF&
Else
    lbl22.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl23_Click()
lbl.Item(22).Caption= cmbNumeros.Text
If lbl.Item(22).Caption= lbl.Item(3).CaptionOr lbl.Item(22).Caption= lbl.Item(4).CaptionOr lbl.Item(22).Caption= lbl.Item(5).CaptionOr lbl.Item(22).Caption= lbl.Item(12).CaptionOr lbl.Item(22).Caption= lbl.Item(13).CaptionOr lbl.Item(22).Caption= lbl.Item(14).CaptionOr lbl.Item(22).Caption= lbl.Item(18).CaptionOr lbl.Item(22).Caption= lbl.Item(19).CaptionOr lbl.Item(22).Caption= lbl.Item(20).CaptionOr lbl.Item(22).Caption= lbl.Item(21).CaptionOr lbl.Item(22).Caption= lbl.Item(23).CaptionOr lbl.Item(22).Caption= lbl.Item(24).CaptionOr lbl.Item(22).Caption= lbl.Item(25).Caption Or lbl.Item(22).Caption= lbl.Item(26).Caption Or lbl.Item(22).Caption= lbl.Item(46).Caption Or lbl.Item(22).Caption= lbl.Item(49).Caption Or lbl.Item(22).Caption= lbl.Item(52).Caption Or lbl.Item(22).Caption= lbl.Item(73).Caption Or lbl.Item(22).Caption= lbl.Item(76).Caption Or lbl.Item(22).Caption= lbl.Item(79).Caption Then
    lbl23.ForeColor = &HFF&
Else
    lbl23.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl24_Click()
lbl.Item(23).Caption= cmbNumeros.Text
If lbl.Item(23).Caption= lbl.Item(3).CaptionOr lbl.Item(23).Caption= lbl.Item(4).CaptionOr lbl.Item(23).Caption= lbl.Item(5).CaptionOr lbl.Item(23).Caption= lbl.Item(12).CaptionOr lbl.Item(23).Caption= lbl.Item(13).CaptionOr lbl.Item(23).Caption= lbl.Item(14).CaptionOr lbl.Item(23).Caption= lbl.Item(18).CaptionOr lbl.Item(23).Caption= lbl.Item(19).CaptionOr lbl.Item(23).Caption= lbl.Item(20).CaptionOr lbl.Item(23).Caption= lbl.Item(21).CaptionOr lbl.Item(23).Caption= lbl.Item(22).CaptionOr lbl.Item(23).Caption= lbl.Item(24).CaptionOr lbl.Item(23).Caption= lbl.Item(25).Caption Or lbl.Item(23).Caption= lbl.Item(26).Caption Or lbl.Item(23).Caption= lbl.Item(47).Caption Or lbl.Item(23).Caption= lbl.Item(50).Caption Or lbl.Item(23).Caption= lbl.Item(53).Caption Or lbl.Item(23).Caption= lbl.Item(74).Caption Or lbl.Item(23).Caption= lbl.Item(77).Caption Or lbl.Item(23).Caption= lbl.Item(80).Caption Then
    lbl24.ForeColor = &HFF&
Else
    lbl24.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl25_Click()
lbl.Item(24).Caption= cmbNumeros.Text
If lbl.Item(24).Caption= lbl.Item(6).CaptionOr lbl.Item(24).Caption= lbl.Item(7).CaptionOr lbl.Item(24).Caption= lbl.Item(8).CaptionOr lbl.Item(24).Caption= lbl.Item(15).CaptionOr lbl.Item(24).Caption= lbl.Item(16).CaptionOr lbl.Item(24).Caption= lbl.Item(17).CaptionOr lbl.Item(24).Caption= lbl.Item(18).CaptionOr lbl.Item(24).Caption= lbl.Item(19).CaptionOr lbl.Item(24).Caption= lbl.Item(20).CaptionOr lbl.Item(24).Caption= lbl.Item(21).CaptionOr lbl.Item(24).Caption= lbl.Item(22).CaptionOr lbl.Item(24).Caption= lbl.Item(23).CaptionOr lbl.Item(24).Caption= lbl.Item(25).Caption Or lbl.Item(24).Caption= lbl.Item(26).Caption Or lbl.Item(24).Caption= lbl.Item(45).Caption Or lbl.Item(24).Caption= lbl.Item(48).Caption Or lbl.Item(24).Caption= lbl.Item(51).Caption Or lbl.Item(24).Caption= lbl.Item(72).Caption Or lbl.Item(24).Caption= lbl.Item(75).Caption Or lbl.Item(24).Caption= lbl.Item(78).Caption Then
    lbl25.ForeColor = &HFF&
Else
    lbl25.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl26_Click()
lbl.Item(25).Caption = cmbNumeros.Text
If lbl.Item(25).Caption = lbl.Item(6).CaptionOr lbl.Item(25).Caption = lbl.Item(7).CaptionOr lbl.Item(25).Caption = lbl.Item(8).CaptionOr lbl.Item(25).Caption = lbl.Item(15).CaptionOr lbl.Item(25).Caption = lbl.Item(16).CaptionOr lbl.Item(25).Caption = lbl.Item(17).CaptionOr lbl.Item(25).Caption = lbl.Item(18).CaptionOr lbl.Item(25).Caption = lbl.Item(19).CaptionOr lbl.Item(25).Caption = lbl.Item(20).CaptionOr lbl.Item(25).Caption = lbl.Item(21).CaptionOr lbl.Item(25).Caption = lbl.Item(22).CaptionOr lbl.Item(25).Caption = lbl.Item(23).CaptionOr lbl.Item(25).Caption = lbl.Item(24).CaptionOr lbl.Item(25).Caption = lbl.Item(26).Caption Or lbl.Item(25).Caption = lbl.Item(46).Caption Or lbl.Item(25).Caption = lbl.Item(49).Caption Or lbl.Item(25).Caption = lbl.Item(52).Caption Or lbl.Item(25).Caption = lbl.Item(73).Caption Or lbl.Item(25).Caption = lbl.Item(76).Caption Or lbl.Item(25).Caption = lbl.Item(79).Caption Then
    lbl26.ForeColor = &HFF&
Else
    lbl26.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl28_Click()
lbl.Item(27).Caption = cmbNumeros.Text
If lbl.Item(27).Caption = lbl.Item(0).CaptionOr lbl.Item(27).Caption = lbl.Item(3).CaptionOr lbl.Item(27).Caption = lbl.Item(6).CaptionOr lbl.Item(27).Caption = lbl.Item(54).Caption Or lbl.Item(27).Caption = lbl.Item(57).Caption Or lbl.Item(27).Caption = lbl.Item(60).Caption Or lbl.Item(27).Caption = lbl.Item(28).Caption Or lbl.Item(27).Caption = lbl.Item(29).Caption Or lbl.Item(27).Caption = lbl.Item(30).Caption Or lbl.Item(27).Caption = lbl.Item(31).Caption Or lbl.Item(27).Caption = lbl.Item(32).Caption Or lbl.Item(27).Caption = lbl.Item(33).Caption Or lbl.Item(27).Caption = lbl.Item(34).Caption Or lbl.Item(27).Caption = lbl.Item(35).Caption Or lbl.Item(27).Caption = lbl.Item(36).Caption Or lbl.Item(27).Caption = lbl.Item(37).Caption Or lbl.Item(27).Caption = lbl.Item(38).Caption Or lbl.Item(27).Caption = lbl.Item(45).Caption Or lbl.Item(27).Caption = lbl.Item(46).Caption Or lbl.Item(27).Caption = lbl.Item(47).Caption Then
    lbl28.ForeColor = &HFF&
Else
    lbl28.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl29_Click()
lbl.Item(28).Caption = cmbNumeros.Text
If lbl.Item(28).Caption = lbl.Item(1).CaptionOr lbl.Item(28).Caption = lbl.Item(4).CaptionOr lbl.Item(28).Caption = lbl.Item(7).CaptionOr lbl.Item(28).Caption = lbl.Item(55).Caption Or lbl.Item(28).Caption = lbl.Item(58).Caption Or lbl.Item(28).Caption = lbl.Item(61).Caption Or lbl.Item(28).Caption = lbl.Item(27).Caption Or lbl.Item(28).Caption = lbl.Item(29).Caption Or lbl.Item(28).Caption = lbl.Item(30).Caption Or lbl.Item(28).Caption = lbl.Item(31).Caption Or lbl.Item(28).Caption = lbl.Item(32).Caption Or lbl.Item(28).Caption = lbl.Item(33).Caption Or lbl.Item(28).Caption = lbl.Item(34).Caption Or lbl.Item(28).Caption = lbl.Item(35).Caption Or lbl.Item(28).Caption = lbl.Item(36).Caption Or lbl.Item(28).Caption = lbl.Item(37).Caption Or lbl.Item(28).Caption = lbl.Item(38).Caption Or lbl.Item(28).Caption = lbl.Item(45).Caption Or lbl.Item(28).Caption = lbl.Item(46).Caption Or lbl.Item(28).Caption = lbl.Item(47).Caption Then
    lbl29.ForeColor = &HFF&
Else
    lbl29.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl3_Click()
lbl.Item(2).Caption= cmbNumeros.Text
If lbl.Item(2).Caption= lbl.Item(10).Caption Or lbl.Item(2).Caption= lbl.Item(10).CaptionOr lbl.Item(2).Caption= lbl.Item(11).CaptionOr lbl.Item(2).Caption= lbl.Item(18).CaptionOr lbl.Item(2).Caption= lbl.Item(19).CaptionOr lbl.Item(2).Caption= lbl.Item(20).CaptionOr lbl.Item(2).Caption= lbl.Item(0).CaptionOr lbl.Item(2).Caption= lbl.Item(1).CaptionOr lbl.Item(2).Caption= lbl.Item(3).CaptionOr lbl.Item(2).Caption= lbl.Item(4).CaptionOr lbl.Item(2).Caption= lbl.Item(5).CaptionOr lbl.Item(2).Caption= lbl.Item(6).CaptionOr lbl.Item(2).Caption= lbl.Item(7).CaptionOr lbl.Item(2).Caption= lbl.Item(8).CaptionOr lbl.Item(2).Caption= lbl.Item(29).Caption Or lbl.Item(2).Caption= lbl.Item(32).Caption Or lbl.Item(2).Caption= lbl.Item(35).Caption Or lbl.Item(2).Caption= lbl.Item(56).Caption Or lbl.Item(2).Caption= lbl.Item(59).Caption Or lbl.Item(2).Caption= lbl.Item(62).Caption Then
    lbl3.ForeColor = &HFF&
Else
    lbl3.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl30_Click()
lbl.Item(29).Caption = cmbNumeros.Text
If lbl.Item(29).Caption = lbl.Item(2).CaptionOr lbl.Item(29).Caption = lbl.Item(5).CaptionOr lbl.Item(29).Caption = lbl.Item(8).CaptionOr lbl.Item(29).Caption = lbl.Item(56).Caption Or lbl.Item(29).Caption = lbl.Item(59).Caption Or lbl.Item(29).Caption = lbl.Item(62).Caption Or lbl.Item(29).Caption = lbl.Item(27).Caption Or lbl.Item(29).Caption = lbl.Item(28).Caption Or lbl.Item(29).Caption = lbl.Item(30).Caption Or lbl.Item(29).Caption = lbl.Item(31).Caption Or lbl.Item(29).Caption = lbl.Item(32).Caption Or lbl.Item(29).Caption = lbl.Item(33).Caption Or lbl.Item(29).Caption = lbl.Item(34).Caption Or lbl.Item(29).Caption = lbl.Item(35).Caption Or lbl.Item(29).Caption = lbl.Item(36).Caption Or lbl.Item(29).Caption = lbl.Item(37).Caption Or lbl.Item(29).Caption = lbl.Item(38).Caption Or lbl.Item(29).Caption = lbl.Item(45).Caption Or lbl.Item(29).Caption = lbl.Item(46).Caption Or lbl.Item(29).Caption = lbl.Item(47).Caption Then
    lbl30.ForeColor = &HFF&
Else
    lbl30.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl31_Click()
lbl.Item(30).Caption = cmbNumeros.Text
If lbl.Item(30).Caption = lbl.Item(0).CaptionOr lbl.Item(30).Caption = lbl.Item(3).CaptionOr lbl.Item(30).Caption = lbl.Item(6).CaptionOr lbl.Item(30).Caption = lbl.Item(54).Caption Or lbl.Item(30).Caption = lbl.Item(57).Caption Or lbl.Item(30).Caption = lbl.Item(60).Caption Or lbl.Item(30).Caption = lbl.Item(27).Caption Or lbl.Item(30).Caption = lbl.Item(28).Caption Or lbl.Item(30).Caption = lbl.Item(29).Caption Or lbl.Item(30).Caption = lbl.Item(31).Caption Or lbl.Item(30).Caption = lbl.Item(32).Caption Or lbl.Item(30).Caption = lbl.Item(33).Caption Or lbl.Item(30).Caption = lbl.Item(34).Caption Or lbl.Item(30).Caption = lbl.Item(35).Caption Or lbl.Item(30).Caption = lbl.Item(39).Caption Or lbl.Item(30).Caption = lbl.Item(40).Caption Or lbl.Item(30).Caption = lbl.Item(41).Caption Or lbl.Item(30).Caption = lbl.Item(48).Caption Or lbl.Item(30).Caption = lbl.Item(49).Caption Or lbl.Item(30).Caption = lbl.Item(50).Caption Then
    lbl31.ForeColor = &HFF&
Else
    lbl31.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl33_Click()
lbl.Item(32).Caption = cmbNumeros.Text
If lbl.Item(32).Caption = lbl.Item(2).CaptionOr lbl.Item(32).Caption = lbl.Item(5).CaptionOr lbl.Item(32).Caption = lbl.Item(8).CaptionOr lbl.Item(32).Caption = lbl.Item(56).Caption Or lbl.Item(32).Caption = lbl.Item(59).Caption Or lbl.Item(32).Caption = lbl.Item(62).Caption Or lbl.Item(32).Caption = lbl.Item(27).Caption Or lbl.Item(32).Caption = lbl.Item(28).Caption Or lbl.Item(32).Caption = lbl.Item(29).Caption Or lbl.Item(32).Caption = lbl.Item(30).Caption Or lbl.Item(32).Caption = lbl.Item(31).Caption Or lbl.Item(32).Caption = lbl.Item(33).Caption Or lbl.Item(32).Caption = lbl.Item(34).Caption Or lbl.Item(32).Caption = lbl.Item(35).Caption Or lbl.Item(32).Caption = lbl.Item(39).Caption Or lbl.Item(32).Caption = lbl.Item(40).Caption Or lbl.Item(32).Caption = lbl.Item(41).Caption Or lbl.Item(32).Caption = lbl.Item(48).Caption Or lbl.Item(32).Caption = lbl.Item(49).Caption Or lbl.Item(32).Caption = lbl.Item(50).Caption Then
    lbl33.ForeColor = &HFF&
Else
    lbl33.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl34_Click()
lbl.Item(33).Caption = cmbNumeros.Text
If lbl.Item(33).Caption = lbl.Item(0).CaptionOr lbl.Item(33).Caption = lbl.Item(3).CaptionOr lbl.Item(33).Caption = lbl.Item(6).CaptionOr lbl.Item(33).Caption = lbl.Item(54).Caption Or lbl.Item(33).Caption = lbl.Item(57).Caption Or lbl.Item(33).Caption = lbl.Item(60).Caption Or lbl.Item(33).Caption = lbl.Item(27).Caption Or lbl.Item(33).Caption = lbl.Item(28).Caption Or lbl.Item(33).Caption = lbl.Item(29).Caption Or lbl.Item(33).Caption = lbl.Item(30).Caption Or lbl.Item(33).Caption = lbl.Item(31).Caption Or lbl.Item(33).Caption = lbl.Item(32).Caption Or lbl.Item(33).Caption = lbl.Item(34).Caption Or lbl.Item(33).Caption = lbl.Item(35).Caption Or lbl.Item(33).Caption = lbl.Item(42).Caption Or lbl.Item(33).Caption = lbl.Item(43).Caption Or lbl.Item(33).Caption = lbl.Item(44).Caption Or lbl.Item(33).Caption = lbl.Item(51).Caption Or lbl.Item(33).Caption = lbl.Item(52).Caption Or lbl.Item(33).Caption = lbl.Item(53).Caption Then
    lbl34.ForeColor = &HFF&
Else
    lbl34.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl35_Click()
lbl.Item(34).Caption = cmbNumeros.Text
If lbl.Item(34).Caption = lbl.Item(1).CaptionOr lbl.Item(34).Caption = lbl.Item(4).CaptionOr lbl.Item(34).Caption = lbl.Item(7).CaptionOr lbl.Item(34).Caption = lbl.Item(55).Caption Or lbl.Item(34).Caption = lbl.Item(58).Caption Or lbl.Item(34).Caption = lbl.Item(61).Caption Or lbl.Item(34).Caption = lbl.Item(27).Caption Or lbl.Item(34).Caption = lbl.Item(28).Caption Or lbl.Item(34).Caption = lbl.Item(29).Caption Or lbl.Item(34).Caption = lbl.Item(30).Caption Or lbl.Item(34).Caption = lbl.Item(31).Caption Or lbl.Item(34).Caption = lbl.Item(32).Caption Or lbl.Item(34).Caption = lbl.Item(33).Caption Or lbl.Item(34).Caption = lbl.Item(35).Caption Or lbl.Item(34).Caption = lbl.Item(42).Caption Or lbl.Item(34).Caption = lbl.Item(43).Caption Or lbl.Item(34).Caption = lbl.Item(44).Caption Or lbl.Item(34).Caption = lbl.Item(51).Caption Or lbl.Item(34).Caption = lbl.Item(52).Caption Or lbl.Item(34).Caption = lbl.Item(53).Caption Then
    lbl35.ForeColor = &HFF&
Else
    lbl35.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl36_Click()
lbl.Item(35).Caption = cmbNumeros.Text
If lbl.Item(35).Caption = lbl.Item(2).CaptionOr lbl.Item(35).Caption = lbl.Item(5).CaptionOr lbl.Item(35).Caption = lbl.Item(8).CaptionOr lbl.Item(35).Caption = lbl.Item(56).Caption Or lbl.Item(35).Caption = lbl.Item(59).Caption Or lbl.Item(35).Caption = lbl.Item(62).Caption Or lbl.Item(35).Caption = lbl.Item(27).Caption Or lbl.Item(35).Caption = lbl.Item(28).Caption Or lbl.Item(35).Caption = lbl.Item(29).Caption Or lbl.Item(35).Caption = lbl.Item(30).Caption Or lbl.Item(35).Caption = lbl.Item(31).Caption Or lbl.Item(35).Caption = lbl.Item(32).Caption Or lbl.Item(35).Caption = lbl.Item(33).Caption Or lbl.Item(35).Caption = lbl.Item(34).Caption Or lbl.Item(35).Caption = lbl.Item(42).Caption Or lbl.Item(35).Caption = lbl.Item(43).Caption Or lbl.Item(35).Caption = lbl.Item(44).Caption Or lbl.Item(35).Caption = lbl.Item(51).Caption Or lbl.Item(35).Caption = lbl.Item(52).Caption Or lbl.Item(35).Caption = lbl.Item(53).Caption Then
    lbl36.ForeColor = &HFF&
Else
    lbl36.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl37_Click()
lbl.Item(36).Caption = cmbNumeros.Text
If lbl.Item(36).Caption = lbl.Item(10).Caption Or lbl.Item(36).Caption = lbl.Item(12).CaptionOr lbl.Item(36).Caption = lbl.Item(15).CaptionOr lbl.Item(36).Caption = lbl.Item(63).Caption Or lbl.Item(36).Caption = lbl.Item(66).Caption Or lbl.Item(36).Caption = lbl.Item(69).Caption Or lbl.Item(36).Caption = lbl.Item(37).Caption Or lbl.Item(36).Caption = lbl.Item(38).Caption Or lbl.Item(36).Caption = lbl.Item(39).Caption Or lbl.Item(36).Caption = lbl.Item(40).Caption Or lbl.Item(36).Caption = lbl.Item(41).Caption Or lbl.Item(36).Caption = lbl.Item(42).Caption Or lbl.Item(36).Caption = lbl.Item(43).Caption Or lbl.Item(36).Caption = lbl.Item(44).Caption Or lbl.Item(36).Caption = lbl.Item(27).Caption Or lbl.Item(36).Caption = lbl.Item(28).Caption Or lbl.Item(36).Caption = lbl.Item(29).Caption Or lbl.Item(36).Caption = lbl.Item(45).Caption Or lbl.Item(36).Caption = lbl.Item(46).Caption Or lbl.Item(36).Caption = lbl.Item(47).Caption Then
    lbl37.ForeColor = &HFF&
Else
    lbl37.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl39_Click()
lbl.Item(38).Caption = cmbNumeros.Text
If lbl.Item(38).Caption = lbl.Item(11).CaptionOr lbl.Item(38).Caption = lbl.Item(14).CaptionOr lbl.Item(38).Caption = lbl.Item(17).CaptionOr lbl.Item(38).Caption = lbl.Item(65).Caption Or lbl.Item(38).Caption = lbl.Item(68).Caption Or lbl.Item(38).Caption = lbl.Item(71).Caption Or lbl.Item(38).Caption = lbl.Item(36).Caption Or lbl.Item(38).Caption = lbl.Item(37).Caption Or lbl.Item(38).Caption = lbl.Item(39).Caption Or lbl.Item(38).Caption = lbl.Item(40).Caption Or lbl.Item(38).Caption = lbl.Item(41).Caption Or lbl.Item(38).Caption = lbl.Item(42).Caption Or lbl.Item(38).Caption = lbl.Item(43).Caption Or lbl.Item(38).Caption = lbl.Item(44).Caption Or lbl.Item(38).Caption = lbl.Item(27).Caption Or lbl.Item(38).Caption = lbl.Item(28).Caption Or lbl.Item(38).Caption = lbl.Item(29).Caption Or lbl.Item(38).Caption = lbl.Item(45).Caption Or lbl.Item(38).Caption = lbl.Item(46).Caption Or lbl.Item(38).Caption = lbl.Item(47).Caption Then
    lbl39.ForeColor = &HFF&
Else
    lbl39.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl4_Click()
lbl.Item(3).Caption= cmbNumeros.Text
If lbl.Item(3).Caption= lbl.Item(12).CaptionOr lbl.Item(3).Caption= lbl.Item(13).CaptionOr lbl.Item(3).Caption= lbl.Item(14).CaptionOr lbl.Item(3).Caption= lbl.Item(21).CaptionOr lbl.Item(3).Caption= lbl.Item(22).CaptionOr lbl.Item(3).Caption= lbl.Item(23).CaptionOr lbl.Item(3).Caption= lbl.Item(0).CaptionOr lbl.Item(3).Caption= lbl.Item(1).CaptionOr lbl.Item(3).Caption= lbl.Item(2).CaptionOr lbl.Item(3).Caption= lbl.Item(4).CaptionOr lbl.Item(3).Caption= lbl.Item(5).CaptionOr lbl.Item(3).Caption= lbl.Item(6).CaptionOr lbl.Item(3).Caption= lbl.Item(7).CaptionOr lbl.Item(3).Caption= lbl.Item(8).CaptionOr lbl.Item(3).Caption= lbl.Item(27).Caption Or lbl.Item(3).Caption= lbl.Item(30).Caption Or lbl.Item(3).Caption= lbl.Item(33).Caption Or lbl.Item(3).Caption= lbl.Item(54).Caption Or lbl.Item(3).Caption= lbl.Item(57).Caption Or lbl.Item(3).Caption= lbl.Item(60).Caption Then
    lbl4.ForeColor = &HFF&
Else
    lbl4.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl40_Click()
lbl.Item(39).Caption = cmbNumeros.Text
If lbl.Item(39).Caption = lbl.Item(10).Caption Or lbl.Item(39).Caption = lbl.Item(12).CaptionOr lbl.Item(39).Caption = lbl.Item(15).CaptionOr lbl.Item(39).Caption = lbl.Item(63).Caption Or lbl.Item(39).Caption = lbl.Item(66).Caption Or lbl.Item(39).Caption = lbl.Item(69).Caption Or lbl.Item(39).Caption = lbl.Item(36).Caption Or lbl.Item(39).Caption = lbl.Item(37).Caption Or lbl.Item(39).Caption = lbl.Item(38).Caption Or lbl.Item(39).Caption = lbl.Item(40).Caption Or lbl.Item(39).Caption = lbl.Item(41).Caption Or lbl.Item(39).Caption = lbl.Item(42).Caption Or lbl.Item(39).Caption = lbl.Item(43).Caption Or lbl.Item(39).Caption = lbl.Item(44).Caption Or lbl.Item(39).Caption = lbl.Item(30).Caption Or lbl.Item(39).Caption = lbl.Item(31).Caption Or lbl.Item(39).Caption = lbl.Item(32).Caption Or lbl.Item(39).Caption = lbl.Item(48).Caption Or lbl.Item(39).Caption = lbl.Item(49).Caption Or lbl.Item(39).Caption = lbl.Item(50).Caption Then
    lbl40.ForeColor = &HFF&
Else
    lbl40.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl41_Click()
lbl.Item(40).Caption = cmbNumeros.Text
If lbl.Item(40).Caption = lbl.Item(10).CaptionOr lbl.Item(40).Caption = lbl.Item(13).CaptionOr lbl.Item(40).Caption = lbl.Item(16).CaptionOr lbl.Item(40).Caption = lbl.Item(64).Caption Or lbl.Item(40).Caption = lbl.Item(67).Caption Or lbl.Item(40).Caption = lbl.Item(70).Caption Or lbl.Item(40).Caption = lbl.Item(36).Caption Or lbl.Item(40).Caption = lbl.Item(37).Caption Or lbl.Item(40).Caption = lbl.Item(38).Caption Or lbl.Item(40).Caption = lbl.Item(39).Caption Or lbl.Item(40).Caption = lbl.Item(41).Caption Or lbl.Item(40).Caption = lbl.Item(42).Caption Or lbl.Item(40).Caption = lbl.Item(43).Caption Or lbl.Item(40).Caption = lbl.Item(44).Caption Or lbl.Item(40).Caption = lbl.Item(30).Caption Or lbl.Item(40).Caption = lbl.Item(31).Caption Or lbl.Item(40).Caption = lbl.Item(32).Caption Or lbl.Item(40).Caption = lbl.Item(48).Caption Or lbl.Item(40).Caption = lbl.Item(49).Caption Or lbl.Item(40).Caption = lbl.Item(50).Caption Then
    lbl41.ForeColor = &HFF&
Else
    lbl41.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl42_Click()
lbl.Item(41).Caption = cmbNumeros.Text
If lbl.Item(41).Caption = lbl.Item(11).CaptionOr lbl.Item(41).Caption = lbl.Item(14).CaptionOr lbl.Item(41).Caption = lbl.Item(17).CaptionOr lbl.Item(41).Caption = lbl.Item(65).Caption Or lbl.Item(41).Caption = lbl.Item(68).Caption Or lbl.Item(41).Caption = lbl.Item(71).Caption Or lbl.Item(41).Caption = lbl.Item(36).Caption Or lbl.Item(41).Caption = lbl.Item(37).Caption Or lbl.Item(41).Caption = lbl.Item(38).Caption Or lbl.Item(41).Caption = lbl.Item(39).Caption Or lbl.Item(41).Caption = lbl.Item(40).Caption Or lbl.Item(41).Caption = lbl.Item(42).Caption Or lbl.Item(41).Caption = lbl.Item(43).Caption Or lbl.Item(41).Caption = lbl.Item(44).Caption Or lbl.Item(41).Caption = lbl.Item(30).Caption Or lbl.Item(41).Caption = lbl.Item(31).Caption Or lbl.Item(41).Caption = lbl.Item(32).Caption Or lbl.Item(41).Caption = lbl.Item(48).Caption Or lbl.Item(41).Caption = lbl.Item(49).Caption Or lbl.Item(41).Caption = lbl.Item(50).Caption Then
    lbl42.ForeColor = &HFF&
Else
    lbl42.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl43_Click()
lbl.Item(42).Caption = cmbNumeros.Text
If lbl.Item(42).Caption = lbl.Item(10).Caption Or lbl.Item(42).Caption = lbl.Item(12).CaptionOr lbl.Item(42).Caption = lbl.Item(15).CaptionOr lbl.Item(42).Caption = lbl.Item(63).Caption Or lbl.Item(42).Caption = lbl.Item(66).Caption Or lbl.Item(42).Caption = lbl.Item(69).Caption Or lbl.Item(42).Caption = lbl.Item(36).Caption Or lbl.Item(42).Caption = lbl.Item(37).Caption Or lbl.Item(42).Caption = lbl.Item(38).Caption Or lbl.Item(42).Caption = lbl.Item(39).Caption Or lbl.Item(42).Caption = lbl.Item(40).Caption Or lbl.Item(42).Caption = lbl.Item(41).Caption Or lbl.Item(42).Caption = lbl.Item(43).Caption Or lbl.Item(42).Caption = lbl.Item(44).Caption Or lbl.Item(42).Caption = lbl.Item(33).Caption Or lbl.Item(42).Caption = lbl.Item(34).Caption Or lbl.Item(42).Caption = lbl.Item(35).Caption Or lbl.Item(42).Caption = lbl.Item(51).Caption Or lbl.Item(42).Caption = lbl.Item(52).Caption Or lbl.Item(42).Caption = lbl.Item(53).Caption Then
    lbl43.ForeColor = &HFF&
Else
    lbl43.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl44_Click()
lbl.Item(43).Caption = cmbNumeros.Text
If lbl.Item(43).Caption = lbl.Item(10).CaptionOr lbl.Item(43).Caption = lbl.Item(13).CaptionOr lbl.Item(43).Caption = lbl.Item(16).CaptionOr lbl.Item(43).Caption = lbl.Item(64).Caption Or lbl.Item(43).Caption = lbl.Item(67).Caption Or lbl.Item(43).Caption = lbl.Item(70).Caption Or lbl.Item(43).Caption = lbl.Item(36).Caption Or lbl.Item(43).Caption = lbl.Item(37).Caption Or lbl.Item(43).Caption = lbl.Item(38).Caption Or lbl.Item(43).Caption = lbl.Item(39).Caption Or lbl.Item(43).Caption = lbl.Item(40).Caption Or lbl.Item(43).Caption = lbl.Item(41).Caption Or lbl.Item(43).Caption = lbl.Item(42).Caption Or lbl.Item(43).Caption = lbl.Item(44).Caption Or lbl.Item(43).Caption = lbl.Item(33).Caption Or lbl.Item(43).Caption = lbl.Item(34).Caption Or lbl.Item(43).Caption = lbl.Item(35).Caption Or lbl.Item(43).Caption = lbl.Item(51).Caption Or lbl.Item(43).Caption = lbl.Item(52).Caption Or lbl.Item(43).Caption = lbl.Item(53).Caption Then
    lbl44.ForeColor = &HFF&
Else
    lbl44.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl45_Click()
lbl.Item(44).Caption = cmbNumeros.Text
If lbl.Item(44).Caption = lbl.Item(11).CaptionOr lbl.Item(44).Caption = lbl.Item(14).CaptionOr lbl.Item(44).Caption = lbl.Item(17).CaptionOr lbl.Item(44).Caption = lbl.Item(65).Caption Or lbl.Item(44).Caption = lbl.Item(68).Caption Or lbl.Item(44).Caption = lbl.Item(71).Caption Or lbl.Item(44).Caption = lbl.Item(36).Caption Or lbl.Item(44).Caption = lbl.Item(37).Caption Or lbl.Item(44).Caption = lbl.Item(38).Caption Or lbl.Item(44).Caption = lbl.Item(39).Caption Or lbl.Item(44).Caption = lbl.Item(40).Caption Or lbl.Item(44).Caption = lbl.Item(41).Caption Or lbl.Item(44).Caption = lbl.Item(42).Caption Or lbl.Item(44).Caption = lbl.Item(43).Caption Or lbl.Item(44).Caption = lbl.Item(33).Caption Or lbl.Item(44).Caption = lbl.Item(34).Caption Or lbl.Item(44).Caption = lbl.Item(35).Caption Or lbl.Item(44).Caption = lbl.Item(51).Caption Or lbl.Item(44).Caption = lbl.Item(52).Caption Or lbl.Item(44).Caption = lbl.Item(53).Caption Then
    lbl45.ForeColor = &HFF&
Else
    lbl45.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl46_Click()
lbl.Item(45).Caption = cmbNumeros.Text
If lbl.Item(45).Caption = lbl.Item(18).CaptionOr lbl.Item(45).Caption = lbl.Item(21).CaptionOr lbl.Item(45).Caption = lbl.Item(24).CaptionOr lbl.Item(45).Caption = lbl.Item(72).Caption Or lbl.Item(45).Caption = lbl.Item(75).Caption Or lbl.Item(45).Caption = lbl.Item(78).Caption Or lbl.Item(45).Caption = lbl.Item(46).Caption Or lbl.Item(45).Caption = lbl.Item(47).Caption Or lbl.Item(45).Caption = lbl.Item(48).Caption Or lbl.Item(45).Caption = lbl.Item(49).Caption Or lbl.Item(45).Caption = lbl.Item(50).Caption Or lbl.Item(45).Caption = lbl.Item(51).Caption Or lbl.Item(45).Caption = lbl.Item(52).Caption Or lbl.Item(45).Caption = lbl.Item(53).Caption Or lbl.Item(45).Caption = lbl.Item(27).Caption Or lbl.Item(45).Caption = lbl.Item(28).Caption Or lbl.Item(45).Caption = lbl.Item(29).Caption Or lbl.Item(45).Caption = lbl.Item(36).Caption Or lbl.Item(45).Caption = lbl.Item(37).Caption Or lbl.Item(45).Caption = lbl.Item(38).Caption Then
    lbl46.ForeColor = &HFF&
Else
    lbl46.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl47_Click()
lbl.Item(46).Caption = cmbNumeros.Text
If lbl.Item(46).Caption = lbl.Item(19).CaptionOr lbl.Item(46).Caption = lbl.Item(22).CaptionOr lbl.Item(46).Caption = lbl.Item(25).Caption Or lbl.Item(46).Caption = lbl.Item(73).Caption Or lbl.Item(46).Caption = lbl.Item(76).Caption Or lbl.Item(46).Caption = lbl.Item(79).Caption Or lbl.Item(46).Caption = lbl.Item(45).Caption Or lbl.Item(46).Caption = lbl.Item(47).Caption Or lbl.Item(46).Caption = lbl.Item(48).Caption Or lbl.Item(46).Caption = lbl.Item(49).Caption Or lbl.Item(46).Caption = lbl.Item(50).Caption Or lbl.Item(46).Caption = lbl.Item(51).Caption Or lbl.Item(46).Caption = lbl.Item(52).Caption Or lbl.Item(46).Caption = lbl.Item(53).Caption Or lbl.Item(46).Caption = lbl.Item(27).Caption Or lbl.Item(46).Caption = lbl.Item(28).Caption Or lbl.Item(46).Caption = lbl.Item(29).Caption Or lbl.Item(46).Caption = lbl.Item(36).Caption Or lbl.Item(46).Caption = lbl.Item(37).Caption Or lbl.Item(46).Caption = lbl.Item(38).Caption Then
    lbl47.ForeColor = &HFF&
Else
    lbl47.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl48_Click()
lbl.Item(47).Caption = cmbNumeros.Text
If lbl.Item(47).Caption = lbl.Item(20).CaptionOr lbl.Item(47).Caption = lbl.Item(23).CaptionOr lbl.Item(47).Caption = lbl.Item(26).Caption Or lbl.Item(47).Caption = lbl.Item(74).Caption Or lbl.Item(47).Caption = lbl.Item(77).Caption Or lbl.Item(47).Caption = lbl.Item(80).Caption Or lbl.Item(47).Caption = lbl.Item(45).Caption Or lbl.Item(47).Caption = lbl.Item(46).Caption Or lbl.Item(47).Caption = lbl.Item(48).Caption Or lbl.Item(47).Caption = lbl.Item(49).Caption Or lbl.Item(47).Caption = lbl.Item(50).Caption Or lbl.Item(47).Caption = lbl.Item(51).Caption Or lbl.Item(47).Caption = lbl.Item(52).Caption Or lbl.Item(47).Caption = lbl.Item(53).Caption Or lbl.Item(47).Caption = lbl.Item(27).Caption Or lbl.Item(47).Caption = lbl.Item(28).Caption Or lbl.Item(47).Caption = lbl.Item(29).Caption Or lbl.Item(47).Caption = lbl.Item(36).Caption Or lbl.Item(47).Caption = lbl.Item(37).Caption Or lbl.Item(47).Caption = lbl.Item(38).Caption Then
    lbl48.ForeColor = &HFF&
Else
    lbl48.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl49_Click()
lbl.Item(48).Caption = cmbNumeros.Text
If lbl.Item(48).Caption = lbl.Item(18).CaptionOr lbl.Item(48).Caption = lbl.Item(21).CaptionOr lbl.Item(48).Caption = lbl.Item(24).CaptionOr lbl.Item(48).Caption = lbl.Item(72).Caption Or lbl.Item(48).Caption = lbl.Item(75).Caption Or lbl.Item(48).Caption = lbl.Item(78).Caption Or lbl.Item(48).Caption = lbl.Item(45).Caption Or lbl.Item(48).Caption = lbl.Item(46).Caption Or lbl.Item(48).Caption = lbl.Item(47).Caption Or lbl.Item(48).Caption = lbl.Item(49).Caption Or lbl.Item(48).Caption = lbl.Item(50).Caption Or lbl.Item(48).Caption = lbl.Item(51).Caption Or lbl.Item(48).Caption = lbl.Item(52).Caption Or lbl.Item(48).Caption = lbl.Item(53).Caption Or lbl.Item(48).Caption = lbl.Item(30).Caption Or lbl.Item(48).Caption = lbl.Item(31).Caption Or lbl.Item(48).Caption = lbl.Item(32).Caption Or lbl.Item(48).Caption = lbl.Item(39).Caption Or lbl.Item(48).Caption = lbl.Item(40).Caption Or lbl.Item(48).Caption = lbl.Item(41).Caption Then
    lbl49.ForeColor = &HFF&
Else
    lbl49.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl5_Click()
lbl.Item(4).Caption= cmbNumeros.Text
If lbl.Item(4).Caption= lbl.Item(28).Caption Or lbl.Item(4).Caption= lbl.Item(31).Caption Or lbl.Item(4).Caption= lbl.Item(34).Caption Or lbl.Item(4).Caption= lbl.Item(55).Caption Or lbl.Item(4).Caption= lbl.Item(58).Caption Or lbl.Item(4).Caption= lbl.Item(61).Caption Or lbl.Item(4).Caption= lbl.Item(0).CaptionOr lbl.Item(4).Caption= lbl.Item(1).CaptionOr lbl.Item(4).Caption= lbl.Item(2).CaptionOr lbl.Item(4).Caption= lbl.Item(3).CaptionOr lbl.Item(4).Caption= lbl.Item(5).CaptionOr lbl.Item(4).Caption= lbl.Item(6).CaptionOr lbl.Item(4).Caption= lbl.Item(7).CaptionOr lbl.Item(4).Caption= lbl.Item(8).CaptionOr lbl.Item(4).Caption= lbl.Item(12).CaptionOr lbl.Item(4).Caption= lbl.Item(13).CaptionOr lbl.Item(4).Caption= lbl.Item(0).CaptionOr lbl.Item(4).Caption= lbl.Item(21).CaptionOr lbl.Item(4).Caption= lbl.Item(22).CaptionOr lbl.Item(4).Caption= lbl.Item(23).CaptionThen
    lbl5.ForeColor = &HFF&
Else
    lbl5.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl50_Click()
lbl.Item(49).Caption = cmbNumeros.Text
If lbl.Item(49).Caption = lbl.Item(19).CaptionOr lbl.Item(49).Caption = lbl.Item(22).CaptionOr lbl.Item(49).Caption = lbl.Item(25).Caption Or lbl.Item(49).Caption = lbl.Item(73).Caption Or lbl.Item(49).Caption = lbl.Item(76).Caption Or lbl.Item(49).Caption = lbl.Item(79).Caption Or lbl.Item(49).Caption = lbl.Item(45).Caption Or lbl.Item(49).Caption = lbl.Item(46).Caption Or lbl.Item(49).Caption = lbl.Item(47).Caption Or lbl.Item(49).Caption = lbl.Item(48).Caption Or lbl.Item(49).Caption = lbl.Item(50).Caption Or lbl.Item(49).Caption = lbl.Item(51).Caption Or lbl.Item(49).Caption = lbl.Item(52).Caption Or lbl.Item(49).Caption = lbl.Item(53).Caption Or lbl.Item(49).Caption = lbl.Item(30).Caption Or lbl.Item(49).Caption = lbl.Item(31).Caption Or lbl.Item(49).Caption = lbl.Item(32).Caption Or lbl.Item(49).Caption = lbl.Item(39).Caption Or lbl.Item(49).Caption = lbl.Item(40).Caption Or lbl.Item(49).Caption = lbl.Item(41).Caption Then
    lbl50.ForeColor = &HFF&
Else
    lbl50.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl51_Click()
lbl.Item(50).Caption = cmbNumeros.Text
If lbl.Item(50).Caption = lbl.Item(20).CaptionOr lbl.Item(50).Caption = lbl.Item(23).CaptionOr lbl.Item(50).Caption = lbl.Item(26).Caption Or lbl.Item(50).Caption = lbl.Item(74).Caption Or lbl.Item(50).Caption = lbl.Item(77).Caption Or lbl.Item(50).Caption = lbl.Item(80).Caption Or lbl.Item(50).Caption = lbl.Item(45).Caption Or lbl.Item(50).Caption = lbl.Item(46).Caption Or lbl.Item(50).Caption = lbl.Item(47).Caption Or lbl.Item(50).Caption = lbl.Item(48).Caption Or lbl.Item(50).Caption = lbl.Item(49).Caption Or lbl.Item(50).Caption = lbl.Item(51).Caption Or lbl.Item(50).Caption = lbl.Item(52).Caption Or lbl.Item(50).Caption = lbl.Item(53).Caption Or lbl.Item(50).Caption = lbl.Item(30).Caption Or lbl.Item(50).Caption = lbl.Item(31).Caption Or lbl.Item(50).Caption = lbl.Item(32).Caption Or lbl.Item(50).Caption = lbl.Item(39).Caption Or lbl.Item(50).Caption = lbl.Item(40).Caption Or lbl.Item(50).Caption = lbl.Item(41).Caption Then
    lbl51.ForeColor = &HFF&
Else
    lbl51.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl53_Click()
lbl.Item(52).Caption = cmbNumeros.Text
If lbl.Item(52).Caption = lbl.Item(19).CaptionOr lbl.Item(52).Caption = lbl.Item(22).CaptionOr lbl.Item(52).Caption = lbl.Item(25).Caption Or lbl.Item(52).Caption = lbl.Item(73).Caption Or lbl.Item(52).Caption = lbl.Item(76).Caption Or lbl.Item(52).Caption = lbl.Item(79).Caption Or lbl.Item(52).Caption = lbl.Item(45).Caption Or lbl.Item(52).Caption = lbl.Item(46).Caption Or lbl.Item(52).Caption = lbl.Item(47).Caption Or lbl.Item(52).Caption = lbl.Item(48).Caption Or lbl.Item(52).Caption = lbl.Item(49).Caption Or lbl.Item(52).Caption = lbl.Item(50).Caption Or lbl.Item(52).Caption = lbl.Item(51).Caption Or lbl.Item(52).Caption = lbl.Item(53).Caption Or lbl.Item(52).Caption = lbl.Item(33).Caption Or lbl.Item(52).Caption = lbl.Item(34).Caption Or lbl.Item(52).Caption = lbl.Item(35).Caption Or lbl.Item(52).Caption = lbl.Item(42).Caption Or lbl.Item(52).Caption = lbl.Item(43).Caption Or lbl.Item(52).Caption = lbl.Item(44).Caption Then
    lbl53.ForeColor = &HFF&
Else
    lbl53.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl54_Click()
lbl.Item(53).Caption = cmbNumeros.Text
If lbl.Item(53).Caption = lbl.Item(20).CaptionOr lbl.Item(53).Caption = lbl.Item(23).CaptionOr lbl.Item(53).Caption = lbl.Item(26).Caption Or lbl.Item(53).Caption = lbl.Item(74).Caption Or lbl.Item(53).Caption = lbl.Item(77).Caption Or lbl.Item(53).Caption = lbl.Item(80).Caption Or lbl.Item(53).Caption = lbl.Item(45).Caption Or lbl.Item(53).Caption = lbl.Item(46).Caption Or lbl.Item(53).Caption = lbl.Item(47).Caption Or lbl.Item(53).Caption = lbl.Item(48).Caption Or lbl.Item(53).Caption = lbl.Item(49).Caption Or lbl.Item(53).Caption = lbl.Item(50).Caption Or lbl.Item(53).Caption = lbl.Item(51).Caption Or lbl.Item(53).Caption = lbl.Item(52).Caption Or lbl.Item(53).Caption = lbl.Item(33).Caption Or lbl.Item(53).Caption = lbl.Item(34).Caption Or lbl.Item(53).Caption = lbl.Item(35).Caption Or lbl.Item(53).Caption = lbl.Item(42).Caption Or lbl.Item(53).Caption = lbl.Item(43).Caption Or lbl.Item(53).Caption = lbl.Item(44).Caption Then
    lbl54.ForeColor = &HFF&
Else
    lbl54.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl55_Click()
lbl.Item(54).Caption = cmbNumeros.Text
If lbl.Item(54).Caption = lbl.Item(0).CaptionOr lbl.Item(54).Caption = lbl.Item(3).CaptionOr lbl.Item(54).Caption = lbl.Item(6).CaptionOr lbl.Item(54).Caption = lbl.Item(27).Caption Or lbl.Item(54).Caption = lbl.Item(30).Caption Or lbl.Item(54).Caption = lbl.Item(33).Caption Or lbl.Item(54).Caption = lbl.Item(55).Caption Or lbl.Item(54).Caption = lbl.Item(56).Caption Or lbl.Item(54).Caption = lbl.Item(57).Caption Or lbl.Item(54).Caption = lbl.Item(58).Caption Or lbl.Item(54).Caption = lbl.Item(59).Caption Or lbl.Item(54).Caption = lbl.Item(60).Caption Or lbl.Item(54).Caption = lbl.Item(61).Caption Or lbl.Item(54).Caption = lbl.Item(62).Caption Or lbl.Item(54).Caption = lbl.Item(63).Caption Or lbl.Item(54).Caption = lbl.Item(64).Caption Or lbl.Item(54).Caption = lbl.Item(65).Caption Or lbl.Item(54).Caption = lbl.Item(72).Caption Or lbl.Item(54).Caption = lbl.Item(73).Caption Or lbl.Item(54).Caption = lbl.Item(74).Caption Then
    lbl55.ForeColor = &HFF&
Else
    lbl55.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl56_Click()
lbl.Item(55).Caption = cmbNumeros.Text
If lbl.Item(55).Caption = lbl.Item(1).CaptionOr lbl.Item(55).Caption = lbl.Item(4).CaptionOr lbl.Item(55).Caption = lbl.Item(7).CaptionOr lbl.Item(55).Caption = lbl.Item(28).Caption Or lbl.Item(55).Caption = lbl.Item(31).Caption Or lbl.Item(55).Caption = lbl.Item(34).Caption Or lbl.Item(55).Caption = lbl.Item(54).Caption Or lbl.Item(55).Caption = lbl.Item(56).Caption Or lbl.Item(55).Caption = lbl.Item(57).Caption Or lbl.Item(55).Caption = lbl.Item(58).Caption Or lbl.Item(55).Caption = lbl.Item(59).Caption Or lbl.Item(55).Caption = lbl.Item(60).Caption Or lbl.Item(55).Caption = lbl.Item(61).Caption Or lbl.Item(55).Caption = lbl.Item(62).Caption Or lbl.Item(55).Caption = lbl.Item(63).Caption Or lbl.Item(55).Caption = lbl.Item(64).Caption Or lbl.Item(55).Caption = lbl.Item(65).Caption Or lbl.Item(55).Caption = lbl.Item(72).Caption Or lbl.Item(55).Caption = lbl.Item(73).Caption Or lbl.Item(55).Caption = lbl.Item(74).Caption Then
    lbl56.ForeColor = &HFF&
Else
    lbl56.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl57_Click()
lbl.Item(56).Caption = cmbNumeros.Text
If lbl.Item(56).Caption = lbl.Item(2).CaptionOr lbl.Item(56).Caption = lbl.Item(5).CaptionOr lbl.Item(56).Caption = lbl.Item(8).CaptionOr lbl.Item(56).Caption = lbl.Item(29).Caption Or lbl.Item(56).Caption = lbl.Item(32).Caption Or lbl.Item(56).Caption = lbl.Item(35).Caption Or lbl.Item(56).Caption = lbl.Item(54).Caption Or lbl.Item(56).Caption = lbl.Item(55).Caption Or lbl.Item(56).Caption = lbl.Item(57).Caption Or lbl.Item(56).Caption = lbl.Item(58).Caption Or lbl.Item(56).Caption = lbl.Item(59).Caption Or lbl.Item(56).Caption = lbl.Item(60).Caption Or lbl.Item(56).Caption = lbl.Item(61).Caption Or lbl.Item(56).Caption = lbl.Item(62).Caption Or lbl.Item(56).Caption = lbl.Item(63).Caption Or lbl.Item(56).Caption = lbl.Item(64).Caption Or lbl.Item(56).Caption = lbl.Item(65).Caption Or lbl.Item(56).Caption = lbl.Item(72).Caption Or lbl.Item(56).Caption = lbl.Item(73).Caption Or lbl.Item(56).Caption = lbl.Item(74).Caption Then
    lbl57.ForeColor = &HFF&
Else
    lbl57.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl58_Click()
lbl.Item(57).Caption = cmbNumeros.Text
If lbl.Item(57).Caption = lbl.Item(0).CaptionOr lbl.Item(57).Caption = lbl.Item(3).CaptionOr lbl.Item(57).Caption = lbl.Item(6).CaptionOr lbl.Item(57).Caption = lbl.Item(27).Caption Or lbl.Item(57).Caption = lbl.Item(30).Caption Or lbl.Item(57).Caption = lbl.Item(33).Caption Or lbl.Item(57).Caption = lbl.Item(54).Caption Or lbl.Item(57).Caption = lbl.Item(55).Caption Or lbl.Item(57).Caption = lbl.Item(56).Caption Or lbl.Item(57).Caption = lbl.Item(58).Caption Or lbl.Item(57).Caption = lbl.Item(59).Caption Or lbl.Item(57).Caption = lbl.Item(60).Caption Or lbl.Item(57).Caption = lbl.Item(61).Caption Or lbl.Item(57).Caption = lbl.Item(62).Caption Or lbl.Item(57).Caption = lbl.Item(66).Caption Or lbl.Item(57).Caption = lbl.Item(67).Caption Or lbl.Item(57).Caption = lbl.Item(68).Caption Or lbl.Item(57).Caption = lbl.Item(75).Caption Or lbl.Item(57).Caption = lbl.Item(76).Caption Or lbl.Item(57).Caption = lbl.Item(77).Caption Then
    lbl58.ForeColor = &HFF&
Else
    lbl58.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl59_Click()
lbl.Item(58).Caption = cmbNumeros.Text
If lbl.Item(58).Caption = lbl.Item(1).CaptionOr lbl.Item(58).Caption = lbl.Item(4).CaptionOr lbl.Item(58).Caption = lbl.Item(7).CaptionOr lbl.Item(58).Caption = lbl.Item(28).Caption Or lbl.Item(58).Caption = lbl.Item(31).Caption Or lbl.Item(58).Caption = lbl.Item(34).Caption Or lbl.Item(58).Caption = lbl.Item(54).Caption Or lbl.Item(58).Caption = lbl.Item(55).Caption Or lbl.Item(58).Caption = lbl.Item(56).Caption Or lbl.Item(58).Caption = lbl.Item(57).Caption Or lbl.Item(58).Caption = lbl.Item(59).Caption Or lbl.Item(58).Caption = lbl.Item(60).Caption Or lbl.Item(58).Caption = lbl.Item(61).Caption Or lbl.Item(58).Caption = lbl.Item(62).Caption Or lbl.Item(58).Caption = lbl.Item(66).Caption Or lbl.Item(58).Caption = lbl.Item(67).Caption Or lbl.Item(58).Caption = lbl.Item(68).Caption Or lbl.Item(58).Caption = lbl.Item(75).Caption Or lbl.Item(58).Caption = lbl.Item(76).Caption Or lbl.Item(58).Caption = lbl.Item(77).Caption Then
    lbl59.ForeColor = &HFF&
Else
    lbl59.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl6_Click()
lbl.Item(5).Caption= cmbNumeros.Text
If lbl.Item(5).Caption= lbl.Item(29).Caption Or lbl.Item(5).Caption= lbl.Item(32).Caption Or lbl.Item(5).Caption= lbl.Item(35).Caption Or lbl.Item(5).Caption= lbl.Item(56).Caption Or lbl.Item(5).Caption= lbl.Item(59).Caption Or lbl.Item(5).Caption= lbl.Item(62).Caption Or lbl.Item(5).Caption= lbl.Item(0).CaptionOr lbl.Item(5).Caption= lbl.Item(1).CaptionOr lbl.Item(5).Caption= lbl.Item(2).CaptionOr lbl.Item(5).Caption= lbl.Item(3).CaptionOr lbl.Item(5).Caption= lbl.Item(4).CaptionOr lbl.Item(5).Caption= lbl.Item(6).CaptionOr lbl.Item(5).Caption= lbl.Item(7).CaptionOr lbl.Item(5).Caption= lbl.Item(8).CaptionOr lbl.Item(5).Caption= lbl.Item(12).CaptionOr lbl.Item(5).Caption= lbl.Item(13).CaptionOr lbl.Item(5).Caption= lbl.Item(14).CaptionOr lbl.Item(5).Caption= lbl.Item(21).CaptionOr lbl.Item(5).Caption= lbl.Item(22).CaptionOr lbl.Item(5).Caption= lbl.Item(23).CaptionThen
    lbl6.ForeColor = &HFF&
Else
    lbl6.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl60_Click()
lbl.Item(59).Caption = cmbNumeros.Text
If lbl.Item(59).Caption = lbl.Item(2).CaptionOr lbl.Item(59).Caption = lbl.Item(5).CaptionOr lbl.Item(59).Caption = lbl.Item(8).CaptionOr lbl.Item(59).Caption = lbl.Item(29).Caption Or lbl.Item(59).Caption = lbl.Item(32).Caption Or lbl.Item(59).Caption = lbl.Item(35).Caption Or lbl.Item(59).Caption = lbl.Item(54).Caption Or lbl.Item(59).Caption = lbl.Item(55).Caption Or lbl.Item(59).Caption = lbl.Item(56).Caption Or lbl.Item(59).Caption = lbl.Item(57).Caption Or lbl.Item(59).Caption = lbl.Item(58).Caption Or lbl.Item(59).Caption = lbl.Item(60).Caption Or lbl.Item(59).Caption = lbl.Item(61).Caption Or lbl.Item(59).Caption = lbl.Item(62).Caption Or lbl.Item(59).Caption = lbl.Item(66).Caption Or lbl.Item(59).Caption = lbl.Item(67).Caption Or lbl.Item(59).Caption = lbl.Item(68).Caption Or lbl.Item(59).Caption = lbl.Item(75).Caption Or lbl.Item(59).Caption = lbl.Item(76).Caption Or lbl.Item(59).Caption = lbl.Item(77).Caption Then
    lbl60.ForeColor = &HFF&
Else
    lbl60.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl62_Click()
lbl.Item(61).Caption = cmbNumeros.Text
If lbl.Item(61).Caption = lbl.Item(1).CaptionOr lbl.Item(61).Caption = lbl.Item(4).CaptionOr lbl.Item(61).Caption = lbl.Item(7).CaptionOr lbl.Item(61).Caption = lbl.Item(28).Caption Or lbl.Item(61).Caption = lbl.Item(31).Caption Or lbl.Item(61).Caption = lbl.Item(34).Caption Or lbl.Item(61).Caption = lbl.Item(54).Caption Or lbl.Item(61).Caption = lbl.Item(55).Caption Or lbl.Item(61).Caption = lbl.Item(56).Caption Or lbl.Item(61).Caption = lbl.Item(57).Caption Or lbl.Item(61).Caption = lbl.Item(58).Caption Or lbl.Item(61).Caption = lbl.Item(59).Caption Or lbl.Item(61).Caption = lbl.Item(60).Caption Or lbl.Item(61).Caption = lbl.Item(62).Caption Or lbl.Item(61).Caption = lbl.Item(69).Caption Or lbl.Item(61).Caption = lbl.Item(70).Caption Or lbl.Item(61).Caption = lbl.Item(71).Caption Or lbl.Item(61).Caption = lbl.Item(78).Caption Or lbl.Item(61).Caption = lbl.Item(79).Caption Or lbl.Item(61).Caption = lbl.Item(80).Caption Then
    lbl62.ForeColor = &HFF&
Else
    lbl62.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl63_Click()
lbl.Item(62).Caption = cmbNumeros.Text
If lbl.Item(62).Caption = lbl.Item(2).CaptionOr lbl.Item(62).Caption = lbl.Item(5).CaptionOr lbl.Item(62).Caption = lbl.Item(8).CaptionOr lbl.Item(62).Caption = lbl.Item(29).Caption Or lbl.Item(62).Caption = lbl.Item(32).Caption Or lbl.Item(62).Caption = lbl.Item(35).Caption Or lbl.Item(62).Caption = lbl.Item(54).Caption Or lbl.Item(62).Caption = lbl.Item(55).Caption Or lbl.Item(62).Caption = lbl.Item(56).Caption Or lbl.Item(62).Caption = lbl.Item(57).Caption Or lbl.Item(62).Caption = lbl.Item(58).Caption Or lbl.Item(62).Caption = lbl.Item(59).Caption Or lbl.Item(62).Caption = lbl.Item(60).Caption Or lbl.Item(62).Caption = lbl.Item(61).Caption Or lbl.Item(62).Caption = lbl.Item(69).Caption Or lbl.Item(62).Caption = lbl.Item(70).Caption Or lbl.Item(62).Caption = lbl.Item(71).Caption Or lbl.Item(62).Caption = lbl.Item(78).Caption Or lbl.Item(62).Caption = lbl.Item(79).Caption Or lbl.Item(62).Caption = lbl.Item(80).Caption Then
    lbl63.ForeColor = &HFF&
Else
    lbl63.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl64_Click()
lbl.Item(63).Caption = cmbNumeros.Text
If lbl.Item(63).Caption = lbl.Item(10).Caption Or lbl.Item(63).Caption = lbl.Item(12).CaptionOr lbl.Item(63).Caption = lbl.Item(15).CaptionOr lbl.Item(63).Caption = lbl.Item(36).Caption Or lbl.Item(63).Caption = lbl.Item(39).Caption Or lbl.Item(63).Caption = lbl.Item(42).Caption Or lbl.Item(63).Caption = lbl.Item(64).Caption Or lbl.Item(63).Caption = lbl.Item(65).Caption Or lbl.Item(63).Caption = lbl.Item(66).Caption Or lbl.Item(63).Caption = lbl.Item(67).Caption Or lbl.Item(63).Caption = lbl.Item(68).Caption Or lbl.Item(63).Caption = lbl.Item(69).Caption Or lbl.Item(63).Caption = lbl.Item(70).Caption Or lbl.Item(63).Caption = lbl.Item(71).Caption Or lbl.Item(63).Caption = lbl.Item(54).Caption Or lbl.Item(63).Caption = lbl.Item(55).Caption Or lbl.Item(63).Caption = lbl.Item(56).Caption Or lbl.Item(63).Caption = lbl.Item(72).Caption Or lbl.Item(63).Caption = lbl.Item(73).Caption Or lbl.Item(63).Caption = lbl.Item(74).Caption Then
    lbl64.ForeColor = &HFF&
Else
    lbl64.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl65_Click()
lbl.Item(64).Caption = cmbNumeros.Text
If lbl.Item(64).Caption = lbl.Item(10).CaptionOr lbl.Item(64).Caption = lbl.Item(13).CaptionOr lbl.Item(64).Caption = lbl.Item(16).CaptionOr lbl.Item(64).Caption = lbl.Item(37).Caption Or lbl.Item(64).Caption = lbl.Item(40).Caption Or lbl.Item(64).Caption = lbl.Item(43).Caption Or lbl.Item(64).Caption = lbl.Item(63).Caption Or lbl.Item(64).Caption = lbl.Item(65).Caption Or lbl.Item(64).Caption = lbl.Item(66).Caption Or lbl.Item(64).Caption = lbl.Item(67).Caption Or lbl.Item(64).Caption = lbl.Item(68).Caption Or lbl.Item(64).Caption = lbl.Item(69).Caption Or lbl.Item(64).Caption = lbl.Item(70).Caption Or lbl.Item(64).Caption = lbl.Item(71).Caption Or lbl.Item(64).Caption = lbl.Item(54).Caption Or lbl.Item(64).Caption = lbl.Item(55).Caption Or lbl.Item(64).Caption = lbl.Item(56).Caption Or lbl.Item(64).Caption = lbl.Item(72).Caption Or lbl.Item(64).Caption = lbl.Item(73).Caption Or lbl.Item(64).Caption = lbl.Item(74).Caption Then
    lbl65.ForeColor = &HFF&
Else
    lbl65.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl66_Click()
lbl.Item(65).Caption = cmbNumeros.Text
If lbl.Item(65).Caption = lbl.Item(11).CaptionOr lbl.Item(65).Caption = lbl.Item(14).CaptionOr lbl.Item(65).Caption = lbl.Item(17).CaptionOr lbl.Item(65).Caption = lbl.Item(38).Caption Or lbl.Item(65).Caption = lbl.Item(41).Caption Or lbl.Item(65).Caption = lbl.Item(44).Caption Or lbl.Item(65).Caption = lbl.Item(63).Caption Or lbl.Item(65).Caption = lbl.Item(64).Caption Or lbl.Item(65).Caption = lbl.Item(66).Caption Or lbl.Item(65).Caption = lbl.Item(67).Caption Or lbl.Item(65).Caption = lbl.Item(68).Caption Or lbl.Item(65).Caption = lbl.Item(69).Caption Or lbl.Item(65).Caption = lbl.Item(70).Caption Or lbl.Item(65).Caption = lbl.Item(71).Caption Or lbl.Item(65).Caption = lbl.Item(54).Caption Or lbl.Item(65).Caption = lbl.Item(55).Caption Or lbl.Item(65).Caption = lbl.Item(56).Caption Or lbl.Item(65).Caption = lbl.Item(72).Caption Or lbl.Item(65).Caption = lbl.Item(73).Caption Or lbl.Item(65).Caption = lbl.Item(74).Caption Then
    lbl66.ForeColor = &HFF&
Else
    lbl66.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl67_Click()
lbl.Item(66).Caption = cmbNumeros.Text
If lbl.Item(66).Caption = lbl.Item(10).Caption Or lbl.Item(66).Caption = lbl.Item(12).CaptionOr lbl.Item(66).Caption = lbl.Item(15).CaptionOr lbl.Item(66).Caption = lbl.Item(36).Caption Or lbl.Item(66).Caption = lbl.Item(39).Caption Or lbl.Item(66).Caption = lbl.Item(42).Caption Or lbl.Item(66).Caption = lbl.Item(63).Caption Or lbl.Item(66).Caption = lbl.Item(64).Caption Or lbl.Item(66).Caption = lbl.Item(65).Caption Or lbl.Item(66).Caption = lbl.Item(67).Caption Or lbl.Item(66).Caption = lbl.Item(68).Caption Or lbl.Item(66).Caption = lbl.Item(69).Caption Or lbl.Item(66).Caption = lbl.Item(70).Caption Or lbl.Item(66).Caption = lbl.Item(71).Caption Or lbl.Item(66).Caption = lbl.Item(57).Caption Or lbl.Item(66).Caption = lbl.Item(58).Caption Or lbl.Item(66).Caption = lbl.Item(59).Caption Or lbl.Item(66).Caption = lbl.Item(75).Caption Or lbl.Item(66).Caption = lbl.Item(76).Caption Or lbl.Item(66).Caption = lbl.Item(77).Caption Then
    lbl67.ForeColor = &HFF&
Else
    lbl67.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl69_Click()
lbl.Item(68).Caption = cmbNumeros.Text
If lbl.Item(68).Caption = lbl.Item(11).CaptionOr lbl.Item(68).Caption = lbl.Item(14).CaptionOr lbl.Item(68).Caption = lbl.Item(17).CaptionOr lbl.Item(68).Caption = lbl.Item(38).Caption Or lbl.Item(68).Caption = lbl.Item(41).Caption Or lbl.Item(68).Caption = lbl.Item(44).Caption Or lbl.Item(68).Caption = lbl.Item(63).Caption Or lbl.Item(68).Caption = lbl.Item(64).Caption Or lbl.Item(68).Caption = lbl.Item(65).Caption Or lbl.Item(68).Caption = lbl.Item(66).Caption Or lbl.Item(68).Caption = lbl.Item(67).Caption Or lbl.Item(68).Caption = lbl.Item(69).Caption Or lbl.Item(68).Caption = lbl.Item(70).Caption Or lbl.Item(68).Caption = lbl.Item(71).Caption Or lbl.Item(68).Caption = lbl.Item(57).Caption Or lbl.Item(68).Caption = lbl.Item(58).Caption Or lbl.Item(68).Caption = lbl.Item(59).Caption Or lbl.Item(68).Caption = lbl.Item(75).Caption Or lbl.Item(68).Caption = lbl.Item(76).Caption Or lbl.Item(68).Caption = lbl.Item(77).Caption Then
    lbl69.ForeColor = &HFF&
Else
    lbl69.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl7_Click()
lbl.Item(6).Caption= cmbNumeros.Text
If lbl.Item(6).Caption= lbl.Item(27).Caption Or lbl.Item(6).Caption= lbl.Item(30).Caption Or lbl.Item(6).Caption= lbl.Item(33).Caption Or lbl.Item(6).Caption= lbl.Item(54).Caption Or lbl.Item(6).Caption= lbl.Item(57).Caption Or lbl.Item(6).Caption= lbl.Item(60).Caption Or lbl.Item(6).Caption= lbl.Item(0).CaptionOr lbl.Item(6).Caption= lbl.Item(1).CaptionOr lbl.Item(6).Caption= lbl.Item(2).CaptionOr lbl.Item(6).Caption= lbl.Item(3).CaptionOr lbl.Item(6).Caption= lbl.Item(4).CaptionOr lbl.Item(6).Caption= lbl.Item(5).CaptionOr lbl.Item(6).Caption= lbl.Item(7).CaptionOr lbl.Item(6).Caption= lbl.Item(8).CaptionOr lbl.Item(6).Caption= lbl.Item(15).CaptionOr lbl.Item(6).Caption= lbl.Item(16).CaptionOr lbl.Item(6).Caption= lbl.Item(17).CaptionOr lbl.Item(6).Caption= lbl.Item(24).CaptionOr lbl.Item(6).Caption= lbl.Item(25).Caption Or lbl.Item(6).Caption= lbl.Item(26).Caption Then
    lbl7.ForeColor = &HFF&
Else
    lbl7.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl70_Click()
lbl.Item(69).Caption = cmbNumeros.Text
If lbl.Item(69).Caption = lbl.Item(10).Caption Or lbl.Item(69).Caption = lbl.Item(12).CaptionOr lbl.Item(69).Caption = lbl.Item(15).CaptionOr lbl.Item(69).Caption = lbl.Item(36).Caption Or lbl.Item(69).Caption = lbl.Item(39).Caption Or lbl.Item(69).Caption = lbl.Item(42).Caption Or lbl.Item(69).Caption = lbl.Item(63).Caption Or lbl.Item(69).Caption = lbl.Item(64).Caption Or lbl.Item(69).Caption = lbl.Item(65).Caption Or lbl.Item(69).Caption = lbl.Item(66).Caption Or lbl.Item(69).Caption = lbl.Item(67).Caption Or lbl.Item(69).Caption = lbl.Item(68).Caption Or lbl.Item(69).Caption = lbl.Item(70).Caption Or lbl.Item(69).Caption = lbl.Item(71).Caption Or lbl.Item(69).Caption = lbl.Item(60).Caption Or lbl.Item(69).Caption = lbl.Item(61).Caption Or lbl.Item(69).Caption = lbl.Item(62).Caption Or lbl.Item(69).Caption = lbl.Item(78).Caption Or lbl.Item(69).Caption = lbl.Item(79).Caption Or lbl.Item(69).Caption = lbl.Item(80).Caption Then
    lbl70.ForeColor = &HFF&
Else
    lbl70.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl71_Click()
lbl.Item(70).Caption = cmbNumeros.Text
If lbl.Item(70).Caption = lbl.Item(10).CaptionOr lbl.Item(70).Caption = lbl.Item(13).CaptionOr lbl.Item(70).Caption = lbl.Item(16).CaptionOr lbl.Item(70).Caption = lbl.Item(37).Caption Or lbl.Item(70).Caption = lbl.Item(40).Caption Or lbl.Item(70).Caption = lbl.Item(43).Caption Or lbl.Item(70).Caption = lbl.Item(63).Caption Or lbl.Item(70).Caption = lbl.Item(64).Caption Or lbl.Item(70).Caption = lbl.Item(65).Caption Or lbl.Item(70).Caption = lbl.Item(66).Caption Or lbl.Item(70).Caption = lbl.Item(67).Caption Or lbl.Item(70).Caption = lbl.Item(68).Caption Or lbl.Item(70).Caption = lbl.Item(69).Caption Or lbl.Item(70).Caption = lbl.Item(71).Caption Or lbl.Item(70).Caption = lbl.Item(60).Caption Or lbl.Item(70).Caption = lbl.Item(61).Caption Or lbl.Item(70).Caption = lbl.Item(62).Caption Or lbl.Item(70).Caption = lbl.Item(78).Caption Or lbl.Item(70).Caption = lbl.Item(79).Caption Or lbl.Item(70).Caption = lbl.Item(80).Caption Then
    lbl71.ForeColor = &HFF&
Else
    lbl71.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl72_Click()
lbl.Item(71).Caption = cmbNumeros.Text
If lbl.Item(71).Caption = lbl.Item(11).CaptionOr lbl.Item(71).Caption = lbl.Item(14).CaptionOr lbl.Item(71).Caption = lbl.Item(17).CaptionOr lbl.Item(71).Caption = lbl.Item(38).Caption Or lbl.Item(71).Caption = lbl.Item(41).Caption Or lbl.Item(71).Caption = lbl.Item(44).Caption Or lbl.Item(71).Caption = lbl.Item(63).Caption Or lbl.Item(71).Caption = lbl.Item(64).Caption Or lbl.Item(71).Caption = lbl.Item(65).Caption Or lbl.Item(71).Caption = lbl.Item(66).Caption Or lbl.Item(71).Caption = lbl.Item(67).Caption Or lbl.Item(71).Caption = lbl.Item(68).Caption Or lbl.Item(71).Caption = lbl.Item(69).Caption Or lbl.Item(71).Caption = lbl.Item(70).Caption Or lbl.Item(71).Caption = lbl.Item(60).Caption Or lbl.Item(71).Caption = lbl.Item(61).Caption Or lbl.Item(71).Caption = lbl.Item(62).Caption Or lbl.Item(71).Caption = lbl.Item(78).Caption Or lbl.Item(71).Caption = lbl.Item(79).Caption Or lbl.Item(71).Caption = lbl.Item(80).Caption Then
    lbl72.ForeColor = &HFF&
Else
    lbl72.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl73_Click()
lbl.Item(72).Caption = cmbNumeros.Text
If lbl.Item(72).Caption = lbl.Item(18).CaptionOr lbl.Item(72).Caption = lbl.Item(21).CaptionOr lbl.Item(72).Caption = lbl.Item(24).CaptionOr lbl.Item(72).Caption = lbl.Item(45).Caption Or lbl.Item(72).Caption = lbl.Item(48).Caption Or lbl.Item(72).Caption = lbl.Item(51).Caption Or lbl.Item(72).Caption = lbl.Item(73).Caption Or lbl.Item(72).Caption = lbl.Item(74).Caption Or lbl.Item(72).Caption = lbl.Item(75).Caption Or lbl.Item(72).Caption = lbl.Item(76).Caption Or lbl.Item(72).Caption = lbl.Item(77).Caption Or lbl.Item(72).Caption = lbl.Item(78).Caption Or lbl.Item(72).Caption = lbl.Item(79).Caption Or lbl.Item(72).Caption = lbl.Item(80).Caption Or lbl.Item(72).Caption = lbl.Item(54).Caption Or lbl.Item(72).Caption = lbl.Item(55).Caption Or lbl.Item(72).Caption = lbl.Item(56).Caption Or lbl.Item(72).Caption = lbl.Item(63).Caption Or lbl.Item(72).Caption = lbl.Item(64).Caption Or lbl.Item(72).Caption = lbl.Item(65).Caption Then
    lbl73.ForeColor = &HFF&
Else
    lbl73.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl74_Click()
lbl.Item(73).Caption = cmbNumeros.Text
If lbl.Item(73).Caption = lbl.Item(19).CaptionOr lbl.Item(73).Caption = lbl.Item(22).CaptionOr lbl.Item(73).Caption = lbl.Item(25).Caption Or lbl.Item(73).Caption = lbl.Item(46).Caption Or lbl.Item(73).Caption = lbl.Item(49).Caption Or lbl.Item(73).Caption = lbl.Item(52).Caption Or lbl.Item(73).Caption = lbl.Item(72).Caption Or lbl.Item(73).Caption = lbl.Item(74).Caption Or lbl.Item(73).Caption = lbl.Item(75).Caption Or lbl.Item(73).Caption = lbl.Item(76).Caption Or lbl.Item(73).Caption = lbl.Item(77).Caption Or lbl.Item(73).Caption = lbl.Item(78).Caption Or lbl.Item(73).Caption = lbl.Item(79).Caption Or lbl.Item(73).Caption = lbl.Item(80).Caption Or lbl.Item(73).Caption = lbl.Item(54).Caption Or lbl.Item(73).Caption = lbl.Item(55).Caption Or lbl.Item(73).Caption = lbl.Item(56).Caption Or lbl.Item(73).Caption = lbl.Item(63).Caption Or lbl.Item(73).Caption = lbl.Item(64).Caption Or lbl.Item(73).Caption = lbl.Item(65).Caption Then
    lbl74.ForeColor = &HFF&
Else
    lbl74.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl76_Click()
lbl.Item(75).Caption = cmbNumeros.Text
If lbl.Item(75).Caption = lbl.Item(18).CaptionOr lbl.Item(75).Caption = lbl.Item(21).CaptionOr lbl.Item(75).Caption = lbl.Item(24).CaptionOr lbl.Item(75).Caption = lbl.Item(45).Caption Or lbl.Item(75).Caption = lbl.Item(48).Caption Or lbl.Item(75).Caption = lbl.Item(51).Caption Or lbl.Item(75).Caption = lbl.Item(72).Caption Or lbl.Item(75).Caption = lbl.Item(73).Caption Or lbl.Item(75).Caption = lbl.Item(74).Caption Or lbl.Item(75).Caption = lbl.Item(76).Caption Or lbl.Item(75).Caption = lbl.Item(77).Caption Or lbl.Item(75).Caption = lbl.Item(78).Caption Or lbl.Item(75).Caption = lbl.Item(79).Caption Or lbl.Item(75).Caption = lbl.Item(80).Caption Or lbl.Item(75).Caption = lbl.Item(57).Caption Or lbl.Item(75).Caption = lbl.Item(58).Caption Or lbl.Item(75).Caption = lbl.Item(59).Caption Or lbl.Item(75).Caption = lbl.Item(66).Caption Or lbl.Item(75).Caption = lbl.Item(67).Caption Or lbl.Item(75).Caption = lbl.Item(68).Caption Then
    lbl76.ForeColor = &HFF&
Else
    lbl76.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl77_Click()
lbl.Item(76).Caption = cmbNumeros.Text
If lbl.Item(76).Caption = lbl.Item(19).CaptionOr lbl.Item(76).Caption = lbl.Item(22).CaptionOr lbl.Item(76).Caption = lbl.Item(25).Caption Or lbl.Item(76).Caption = lbl.Item(46).Caption Or lbl.Item(76).Caption = lbl.Item(49).Caption Or lbl.Item(76).Caption = lbl.Item(52).Caption Or lbl.Item(76).Caption = lbl.Item(72).Caption Or lbl.Item(76).Caption = lbl.Item(73).Caption Or lbl.Item(76).Caption = lbl.Item(74).Caption Or lbl.Item(76).Caption = lbl.Item(75).Caption Or lbl.Item(76).Caption = lbl.Item(77).Caption Or lbl.Item(76).Caption = lbl.Item(78).Caption Or lbl.Item(76).Caption = lbl.Item(79).Caption Or lbl.Item(76).Caption = lbl.Item(80).Caption Or lbl.Item(76).Caption = lbl.Item(57).Caption Or lbl.Item(76).Caption = lbl.Item(58).Caption Or lbl.Item(76).Caption = lbl.Item(59).Caption Or lbl.Item(76).Caption = lbl.Item(66).Caption Or lbl.Item(76).Caption = lbl.Item(67).Caption Or lbl.Item(76).Caption = lbl.Item(68).Caption Then
    lbl77.ForeColor = &HFF&
Else
    lbl77.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl78_Click()
lbl.Item(77).Caption = cmbNumeros.Text
If lbl.Item(77).Caption = lbl.Item(20).CaptionOr lbl.Item(77).Caption = lbl.Item(23).CaptionOr lbl.Item(77).Caption = lbl.Item(26).Caption Or lbl.Item(77).Caption = lbl.Item(47).Caption Or lbl.Item(77).Caption = lbl.Item(50).Caption Or lbl.Item(77).Caption = lbl.Item(53).Caption Or lbl.Item(77).Caption = lbl.Item(72).Caption Or lbl.Item(77).Caption = lbl.Item(73).Caption Or lbl.Item(77).Caption = lbl.Item(74).Caption Or lbl.Item(77).Caption = lbl.Item(75).Caption Or lbl.Item(77).Caption = lbl.Item(76).Caption Or lbl.Item(77).Caption = lbl.Item(78).Caption Or lbl.Item(77).Caption = lbl.Item(79).Caption Or lbl.Item(77).Caption = lbl.Item(80).Caption Or lbl.Item(77).Caption = lbl.Item(57).Caption Or lbl.Item(77).Caption = lbl.Item(58).Caption Or lbl.Item(77).Caption = lbl.Item(59).Caption Or lbl.Item(77).Caption = lbl.Item(66).Caption Or lbl.Item(77).Caption = lbl.Item(67).Caption Or lbl.Item(77).Caption = lbl.Item(68).Caption Then
    lbl78.ForeColor = &HFF&
Else
    lbl78.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl79_Click()
lbl.Item(78).Caption = cmbNumeros.Text
If lbl.Item(78).Caption = lbl.Item(18).CaptionOr lbl.Item(78).Caption = lbl.Item(21).CaptionOr lbl.Item(78).Caption = lbl.Item(24).CaptionOr lbl.Item(78).Caption = lbl.Item(45).Caption Or lbl.Item(78).Caption = lbl.Item(48).Caption Or lbl.Item(78).Caption = lbl.Item(51).Caption Or lbl.Item(78).Caption = lbl.Item(72).Caption Or lbl.Item(78).Caption = lbl.Item(73).Caption Or lbl.Item(78).Caption = lbl.Item(74).Caption Or lbl.Item(78).Caption = lbl.Item(75).Caption Or lbl.Item(78).Caption = lbl.Item(76).Caption Or lbl.Item(78).Caption = lbl.Item(77).Caption Or lbl.Item(78).Caption = lbl.Item(79).Caption Or lbl.Item(78).Caption = lbl.Item(80).Caption Or lbl.Item(78).Caption = lbl.Item(60).Caption Or lbl.Item(78).Caption = lbl.Item(61).Caption Or lbl.Item(78).Caption = lbl.Item(62).Caption Or lbl.Item(78).Caption = lbl.Item(69).Caption Or lbl.Item(78).Caption = lbl.Item(70).Caption Or lbl.Item(78).Caption = lbl.Item(71).Caption Then
    lbl79.ForeColor = &HFF&
Else
    lbl79.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl8_Click()
lbl.Item(7).Caption= cmbNumeros.Text
If lbl.Item(7).Caption= lbl.Item(28).Caption Or lbl.Item(7).Caption= lbl.Item(31).Caption Or lbl.Item(7).Caption= lbl.Item(34).Caption Or lbl.Item(7).Caption= lbl.Item(55).Caption Or lbl.Item(7).Caption= lbl.Item(58).Caption Or lbl.Item(7).Caption= lbl.Item(61).Caption Or lbl.Item(7).Caption= lbl.Item(0).CaptionOr lbl.Item(7).Caption= lbl.Item(1).CaptionOr lbl.Item(7).Caption= lbl.Item(2).CaptionOr lbl.Item(7).Caption= lbl.Item(3).CaptionOr lbl.Item(7).Caption= lbl.Item(4).CaptionOr lbl.Item(7).Caption= lbl.Item(5).CaptionOr lbl.Item(7).Caption= lbl.Item(6).CaptionOr lbl.Item(7).Caption= lbl.Item(8).CaptionOr lbl.Item(7).Caption= lbl.Item(15).CaptionOr lbl.Item(7).Caption= lbl.Item(16).CaptionOr lbl.Item(7).Caption= lbl.Item(17).CaptionOr lbl.Item(7).Caption= lbl.Item(24).CaptionOr lbl.Item(7).Caption= lbl.Item(25).Caption Or lbl.Item(7).Caption= lbl.Item(26).Caption Then
    lbl8.ForeColor = &HFF&
Else
    lbl8.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl80_Click()
lbl.Item(79).Caption = cmbNumeros.Text
If lbl.Item(79).Caption = lbl.Item(19).CaptionOr lbl.Item(79).Caption = lbl.Item(22).CaptionOr lbl.Item(79).Caption = lbl.Item(25).Caption Or lbl.Item(79).Caption = lbl.Item(46).Caption Or lbl.Item(79).Caption = lbl.Item(49).Caption Or lbl.Item(79).Caption = lbl.Item(52).Caption Or lbl.Item(79).Caption = lbl.Item(72).Caption Or lbl.Item(79).Caption = lbl.Item(73).Caption Or lbl.Item(79).Caption = lbl.Item(74).Caption Or lbl.Item(79).Caption = lbl.Item(75).Caption Or lbl.Item(79).Caption = lbl.Item(76).Caption Or lbl.Item(79).Caption = lbl.Item(77).Caption Or lbl.Item(79).Caption = lbl.Item(78).Caption Or lbl.Item(79).Caption = lbl.Item(80).Caption Or lbl.Item(79).Caption = lbl.Item(60).Caption Or lbl.Item(79).Caption = lbl.Item(61).Caption Or lbl.Item(79).Caption = lbl.Item(62).Caption Or lbl.Item(79).Caption = lbl.Item(69).Caption Or lbl.Item(79).Caption = lbl.Item(70).Caption Or lbl.Item(79).Caption = lbl.Item(71).Caption Then
    lbl80.ForeColor = &HFF&
Else
    lbl80.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl81_Click()
lbl.Item(80).Caption = cmbNumeros.Text
If lbl.Item(80).Caption = lbl.Item(20).CaptionOr lbl.Item(80).Caption = lbl.Item(23).CaptionOr lbl.Item(80).Caption = lbl.Item(26).Caption Or lbl.Item(80).Caption = lbl.Item(47).Caption Or lbl.Item(80).Caption = lbl.Item(50).Caption Or lbl.Item(80).Caption = lbl.Item(53).Caption Or lbl.Item(80).Caption = lbl.Item(72).Caption Or lbl.Item(80).Caption = lbl.Item(73).Caption Or lbl.Item(80).Caption = lbl.Item(74).Caption Or lbl.Item(80).Caption = lbl.Item(75).Caption Or lbl.Item(80).Caption = lbl.Item(76).Caption Or lbl.Item(80).Caption = lbl.Item(77).Caption Or lbl.Item(80).Caption = lbl.Item(78).Caption Or lbl.Item(80).Caption = lbl.Item(79).Caption Or lbl.Item(80).Caption = lbl.Item(60).Caption Or lbl.Item(80).Caption = lbl.Item(61).Caption Or lbl.Item(80).Caption = lbl.Item(62).Caption Or lbl.Item(80).Caption = lbl.Item(69).Caption Or lbl.Item(80).Caption = lbl.Item(70).Caption Or lbl.Item(80).Caption = lbl.Item(71).Caption Then
    lbl81.ForeColor = &HFF&
Else
    lbl81.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl9_Click()
lbl.Item(8).Caption= cmbNumeros.Text
If lbl.Item(8).Caption= lbl.Item(29).Caption Or lbl.Item(8).Caption= lbl.Item(32).Caption Or lbl.Item(8).Caption= lbl.Item(35).Caption Or lbl.Item(8).Caption= lbl.Item(56).Caption Or lbl.Item(8).Caption= lbl.Item(59).Caption Or lbl.Item(8).Caption= lbl.Item(62).Caption Or lbl.Item(8).Caption= lbl.Item(0).CaptionOr lbl.Item(8).Caption= lbl.Item(1).CaptionOr lbl.Item(8).Caption= lbl.Item(2).CaptionOr lbl.Item(8).Caption= lbl.Item(3).CaptionOr lbl.Item(8).Caption= lbl.Item(4).CaptionOr lbl.Item(8).Caption= lbl.Item(5).CaptionOr lbl.Item(8).Caption= lbl.Item(6).CaptionOr lbl.Item(8).Caption= lbl.Item(7).CaptionOr lbl.Item(8).Caption= lbl.Item(15).CaptionOr lbl.Item(8).Caption= lbl.Item(16).CaptionOr lbl.Item(8).Caption= lbl.Item(17).CaptionOr lbl.Item(8).Caption= lbl.Item(24).CaptionOr lbl.Item(8).Caption= lbl.Item(25).Caption Or lbl.Item(8).Caption= lbl.Item(26).Caption Then
    lbl9.ForeColor = &HFF&
Else
    lbl9.ForeColor = &H0&
End If
    Call Ganhou
End Sub

    End If
End Sub
Sub Limpar()
For A = 0 To 80
lbl.Item(A).Caption = ""
Next
End Sub
