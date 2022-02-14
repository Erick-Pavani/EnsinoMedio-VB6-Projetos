VERSION 5.00
Begin VB.Form frmSudoku 
   BackColor       =   &H000000FF&
   Caption         =   "Sudoku"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10950
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
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
      Height          =   975
      Left            =   7800
      TabIndex        =   92
      Top             =   6000
      Width           =   2775
   End
   Begin VB.Frame fra3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   5040
      TabIndex        =   82
      Top             =   240
      Width           =   2175
      Begin VB.Label lbl27 
         BackColor       =   &H00FF0000&
         Caption         =   " 7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   1560
         TabIndex        =   91
         Tag             =   "x"
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl24 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   90
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl21 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   89
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl26 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   88
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl23 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   87
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl20 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   86
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl25 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   85
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl22 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   84
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl19 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   83
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame fra6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   5040
      TabIndex        =   72
      Top             =   2520
      Width           =   2175
      Begin VB.Label lbl54 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   81
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl51 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   80
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl48 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   79
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl53 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   78
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl50 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   77
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl47 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   76
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl52 
         BackColor       =   &H00FF0000&
         Caption         =   " 8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   120
         TabIndex        =   75
         Tag             =   "x"
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl49 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   74
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl46 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   73
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame fra9 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   5040
      TabIndex        =   62
      Top             =   4800
      Width           =   2175
      Begin VB.Label lbl81 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   71
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl78 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   70
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl75 
         BackColor       =   &H00FF0000&
         Caption         =   " 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   1560
         TabIndex        =   69
         Tag             =   "x"
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl80 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   68
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl77 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   67
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl74 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   66
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl79 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   65
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl76 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   64
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl73 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   63
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame fra8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   2760
      TabIndex        =   52
      Top             =   4800
      Width           =   2295
      Begin VB.Label lbl72 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   61
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl69 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   60
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl66 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   59
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl71 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   58
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl68 
         BackColor       =   &H00FF0000&
         Caption         =   " 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   840
         TabIndex        =   57
         Tag             =   "x"
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl65 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   56
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl70 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   55
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl67 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   54
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl64 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   53
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame fra5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   2760
      TabIndex        =   42
      Top             =   2520
      Width           =   2295
      Begin VB.Label lbl45 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   51
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl42 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   50
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl39 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   49
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl44 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   48
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl41 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   47
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl38 
         BackColor       =   &H00FF0000&
         Caption         =   " 6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   840
         TabIndex        =   46
         Tag             =   "x"
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl43 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   45
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl40 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl37 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame fra2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   2760
      TabIndex        =   32
      Top             =   240
      Width           =   2295
      Begin VB.Label lbl10 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl13 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   40
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl16 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl11 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   38
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl14 
         BackColor       =   &H00FF0000&
         Caption         =   " 9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   840
         TabIndex        =   37
         Tag             =   "x"
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl17 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   36
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl12 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   35
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl15 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   34
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl18 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   33
         Top             =   1560
         Width           =   495
      End
   End
   Begin VB.Frame fra1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   480
      TabIndex        =   22
      Top             =   240
      Width           =   2295
      Begin VB.Label lbl9 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   31
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   30
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   29
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl8 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   28
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   27
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl2 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   26
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl7 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl4 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl1 
         BackColor       =   &H00FF0000&
         Caption         =   " 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Tag             =   "x"
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame fra7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   480
      TabIndex        =   12
      Top             =   4800
      Width           =   2295
      Begin VB.Label lbl55 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl58 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl61 
         BackColor       =   &H00FF0000&
         Caption         =   " 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Tag             =   "x"
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl56 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   18
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl59 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   17
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl62 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   16
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl57 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   15
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl60 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   14
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl63 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   13
         Top             =   1560
         Width           =   495
      End
   End
   Begin VB.Frame fra4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
      Begin VB.Label lbl36 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   11
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl33 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl30 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   1560
         TabIndex        =   9
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl35 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   8
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl32 
         BackColor       =   &H00FF0000&
         Caption         =   " 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Tag             =   "x"
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl29 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   840
         TabIndex        =   6
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl34 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbl31 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl28 
         BackColor       =   &H00FF0000&
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
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbNumeros 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   8400
      List            =   "Form1.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4080
      Width           =   1455
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
      Height          =   855
      Left            =   7680
      TabIndex        =   1
      Tag             =   "x"
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "frmSudoku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Ganhou()
If lbl2.ForeColor = &H0& And lbl3.ForeColor = &H0& And lbl4.ForeColor = &H0& And lbl5.ForeColor = &H0& And lbl6.ForeColor = &H0& And lbl7.ForeColor = &H0& And lbl8.ForeColor = &H0& And lbl9.ForeColor = &H0& And lbl10.ForeColor = &H0& And lbl11.ForeColor = &H0& And lbl12.ForeColor = &H0& And lbl13.ForeColor = &H0& And lbl15.ForeColor = &H0& And lbl16.ForeColor = &H0& And lbl17.ForeColor = &H0& And lbl18.ForeColor = &H0& And lbl19.ForeColor = &H0& And lbl20.ForeColor = &H0& And lbl21.ForeColor = &H0& And lbl22.ForeColor = &H0& And lbl23.ForeColor = &H0& And lbl24.ForeColor = &H0& And lbl25.ForeColor = &H0& And lbl26.ForeColor = &H0& And lbl28.ForeColor = &H0& And lbl29.ForeColor = &H0& And lbl30.ForeColor = &H0& And lbl31.ForeColor = &H0& And lbl33.ForeColor = &H0& And lbl34.ForeColor = &H0& _
    And lbl35.ForeColor = &H0& And lbl36.ForeColor = &H0& And lbl37.ForeColor = &H0& And lbl39.ForeColor = &H0& And lbl40.ForeColor = &H0& And lbl41.ForeColor = &H0& And lbl42.ForeColor = &H0& And lbl43.ForeColor = &H0& And lbl44.ForeColor = &H0& And lbl45.ForeColor = &H0& And lbl46.ForeColor = &H0& And lbl47.ForeColor = &H0& And lbl48.ForeColor = &H0& And lbl49.ForeColor = &H0& And lbl50.ForeColor = &H0& And lbl51.ForeColor = &H0& And lbl53.ForeColor = &H0& And lbl54.ForeColor = &H0& And lbl55.ForeColor = &H0& And lbl56.ForeColor = &H0& And lbl57.ForeColor = &H0& And lbl58.ForeColor = &H0& And lbl59.ForeColor = &H0& And lbl60.ForeColor = &H0& And lbl62.ForeColor = &H0& And lbl63.ForeColor = &H0& And lbl64.ForeColor = &H0& And lbl65.ForeColor = &H0& And lbl66.ForeColor = &H0& _
    And lbl67.ForeColor = &H0& And lbl69.ForeColor = &H0& And lbl70.ForeColor = &H0& And lbl71.ForeColor = &H0& And lbl72.ForeColor = &H0& And lbl73.ForeColor = &H0& And lbl74.ForeColor = &H0& And lbl76.ForeColor = &H0& And lbl77.ForeColor = &H0& And lbl78.ForeColor = &H0& And lbl79.ForeColor = &H0& And lbl80.ForeColor = &H0& And lbl81.ForeColor = &H0& And Not lbl2.Caption = "" And Not lbl3.Caption = "" And Not lbl4.Caption = "" And Not lbl5.Caption = "" And Not lbl6.Caption = "" And Not lbl7.Caption = "" And Not lbl8.Caption = "" And Not lbl9.Caption = "" And Not lbl10.Caption = "" And Not lbl11.Caption = "" And Not lbl12.Caption = "" And Not lbl13.Caption = "" And Not lbl15.Caption = "" And Not lbl16.Caption = "" And Not lbl17.Caption = "" And Not lbl18.Caption = "" And Not lbl19.Caption = "" And Not lbl20.Caption = "" And Not lbl21.Caption = "" And Not lbl22.Caption = "" And Not lbl23.Caption = "" And Not lbl24.Caption = "" _
    And Not lbl25.Caption = "" And Not lbl26.Caption = "" And Not lbl28.Caption = "" And Not lbl29.Caption = "" And Not lbl30.Caption = "" And Not lbl31.Caption = "" And Not lbl33.Caption = "" And Not lbl34.Caption = "" And Not lbl35.Caption = "" And Not lbl36.Caption = "" And Not lbl37.Caption = "" And Not lbl39.Caption = "" And Not lbl40.Caption = "" And Not lbl41.Caption = "" And Not lbl42.Caption = "" And Not lbl43.Caption = "" And Not lbl44.Caption = "" And Not lbl45.Caption = "" And Not lbl46.Caption = "" And Not lbl47.Caption = "" And Not lbl48.Caption = "" And Not lbl49.Caption = "" And Not lbl50.Caption = "" And Not lbl51.Caption = "" And Not lbl53.Caption = "" And Not lbl54.Caption = "" And Not lbl55.Caption = "" And Not lbl56.Caption = "" And Not lbl57.Caption = "" And Not lbl58.Caption = "" And Not lbl59.Caption = "" And Not lbl60.Caption = "" And Not lbl62.Caption = "" And Not lbl63.Caption = "" And Not lbl64.Caption = "" And Not lbl65.Caption = "" And Not lbl66.Caption = "" _
    And Not lbl67.Caption = "" And Not lbl69.Caption = "" And Not lbl70.Caption = "" And Not lbl71.Caption = "" And Not lbl72.Caption = "" And Not lbl73.Caption = "" And Not lbl74.Caption = "" And Not lbl76.Caption = "" And Not lbl77.Caption = "" And Not lbl78.Caption = "" And Not lbl79.Caption = "" And Not lbl80.Caption = "" And Not lbl81.Caption = "" Then
    frmGanhou.Show
    Me.Hide
    lbl2.Caption = ""
    lbl3.Caption = ""
    lbl4.Caption = ""
    lbl5.Caption = ""
    lbl6.Caption = ""
    lbl7.Caption = ""
    lbl8.Caption = ""
    lbl9.Caption = ""
    lbl10.Caption = ""
    lbl11.Caption = ""
    lbl12.Caption = ""
    lbl13.Caption = ""
    lbl15.Caption = ""
    lbl16.Caption = ""
    lbl17.Caption = ""
    lbl18.Caption = ""
    lbl19.Caption = ""
    lbl20.Caption = ""
    lbl21.Caption = ""
    lbl22.Caption = ""
    lbl23.Caption = ""
    lbl24.Caption = ""
    lbl25.Caption = ""
    lbl26.Caption = ""
    lbl28.Caption = ""
    lbl29.Caption = ""
    lbl30.Caption = ""
    lbl31.Caption = ""
    lbl33.Caption = ""
    lbl34.Caption = ""
    lbl35.Caption = ""
    lbl36.Caption = ""
    lbl37.Caption = ""
    lbl39.Caption = ""
    lbl40.Caption = ""
    lbl41.Caption = ""
    lbl42.Caption = ""
    lbl43.Caption = ""
    lbl44.Caption = ""
    lbl45.Caption = ""
    lbl46.Caption = ""
    lbl47.Caption = ""
    lbl48.Caption = ""
    lbl49.Caption = ""
    lbl50.Caption = ""
    lbl51.Caption = ""
    lbl53.Caption = ""
    lbl54.Caption = ""
    lbl55.Caption = ""
    lbl56.Caption = ""
    lbl57.Caption = ""
    lbl58.Caption = ""
    lbl59.Caption = ""
    lbl60.Caption = ""
    lbl62.Caption = ""
    lbl63.Caption = ""
    lbl64.Caption = ""
    lbl65.Caption = ""
    lbl66.Caption = ""
    lbl67.Caption = ""
    lbl69.Caption = ""
    lbl70.Caption = ""
    lbl71.Caption = ""
    lbl72.Caption = ""
    lbl73.Caption = ""
    lbl74.Caption = ""
    lbl76.Caption = ""
    lbl77.Caption = ""
    lbl78.Caption = ""
    lbl79.Caption = ""
    lbl80.Caption = ""
    lbl81.Caption = ""
End If
End Sub
Private Sub cmdLimpar_Click()
lbl2.Caption = ""
lbl3.Caption = ""
lbl4.Caption = ""
lbl5.Caption = ""
lbl6.Caption = ""
lbl7.Caption = ""
lbl8.Caption = ""
lbl9.Caption = ""
lbl10.Caption = ""
lbl11.Caption = ""
lbl12.Caption = ""
lbl13.Caption = ""
lbl15.Caption = ""
lbl16.Caption = ""
lbl17.Caption = ""
lbl18.Caption = ""
lbl19.Caption = ""
lbl20.Caption = ""
lbl21.Caption = ""
lbl22.Caption = ""
lbl23.Caption = ""
lbl24.Caption = ""
lbl25.Caption = ""
lbl26.Caption = ""
lbl28.Caption = ""
lbl29.Caption = ""
lbl30.Caption = ""
lbl31.Caption = ""
lbl33.Caption = ""
lbl34.Caption = ""
lbl35.Caption = ""
lbl36.Caption = ""
lbl37.Caption = ""
lbl39.Caption = ""
lbl40.Caption = ""
lbl41.Caption = ""
lbl42.Caption = ""
lbl43.Caption = ""
lbl44.Caption = ""
lbl45.Caption = ""
lbl46.Caption = ""
lbl47.Caption = ""
lbl48.Caption = ""
lbl49.Caption = ""
lbl50.Caption = ""
lbl51.Caption = ""
lbl53.Caption = ""
lbl54.Caption = ""
lbl55.Caption = ""
lbl56.Caption = ""
lbl57.Caption = ""
lbl58.Caption = ""
lbl59.Caption = ""
lbl60.Caption = ""
lbl62.Caption = ""
lbl63.Caption = ""
lbl64.Caption = ""
lbl65.Caption = ""
lbl66.Caption = ""
lbl67.Caption = ""
lbl69.Caption = ""
lbl70.Caption = ""
lbl71.Caption = ""
lbl72.Caption = ""
lbl73.Caption = ""
lbl74.Caption = ""
lbl76.Caption = ""
lbl77.Caption = ""
lbl78.Caption = ""
lbl79.Caption = ""
lbl80.Caption = ""
lbl81.Caption = ""
lbl2.ForeColor = &H0&
lbl3.ForeColor = &H0&
lbl4.ForeColor = &H0&
lbl5.ForeColor = &H0&
lbl6.ForeColor = &H0&
lbl7.ForeColor = &H0&
lbl8.ForeColor = &H0&
lbl9.ForeColor = &H0&
lbl10.ForeColor = &H0&
lbl11.ForeColor = &H0&
lbl12.ForeColor = &H0&
lbl13.ForeColor = &H0&
lbl15.ForeColor = &H0&
lbl16.ForeColor = &H0&
lbl17.ForeColor = &H0&
lbl18.ForeColor = &H0&
lbl19.ForeColor = &H0&
lbl20.ForeColor = &H0&
lbl21.ForeColor = &H0&
lbl22.ForeColor = &H0&
lbl23.ForeColor = &H0&
lbl24.ForeColor = &H0&
lbl25.ForeColor = &H0&
lbl26.ForeColor = &H0&
lbl28.ForeColor = &H0&
lbl29.ForeColor = &H0&
lbl30.ForeColor = &H0&
lbl31.ForeColor = &H0&
lbl33.ForeColor = &H0&
lbl34.ForeColor = &H0&
lbl35.ForeColor = &H0&
lbl36.ForeColor = &H0&
lbl37.ForeColor = &H0&
lbl39.ForeColor = &H0&
lbl40.ForeColor = &H0&
lbl41.ForeColor = &H0&
lbl42.ForeColor = &H0&
lbl42.ForeColor = &H0&
lbl43.ForeColor = &H0&
lbl44.ForeColor = &H0&
lbl45.ForeColor = &H0&
lbl46.ForeColor = &H0&
lbl47.ForeColor = &H0&
lbl48.ForeColor = &H0&
lbl49.ForeColor = &H0&
lbl50.ForeColor = &H0&
lbl51.ForeColor = &H0&
lbl53.ForeColor = &H0&
lbl54.ForeColor = &H0&
lbl55.ForeColor = &H0&
lbl56.ForeColor = &H0&
lbl57.ForeColor = &H0&
lbl58.ForeColor = &H0&
lbl59.ForeColor = &H0&
lbl60.ForeColor = &H0&
lbl62.ForeColor = &H0&
lbl63.ForeColor = &H0&
lbl64.ForeColor = &H0&
lbl65.ForeColor = &H0&
lbl66.ForeColor = &H0&
lbl67.ForeColor = &H0&
lbl69.ForeColor = &H0&
lbl70.ForeColor = &H0&
lbl71.ForeColor = &H0&
lbl72.ForeColor = &H0&
lbl73.ForeColor = &H0&
lbl74.ForeColor = &H0&
lbl76.ForeColor = &H0&
lbl77.ForeColor = &H0&
lbl78.ForeColor = &H0&
lbl79.ForeColor = &H0&
lbl80.ForeColor = &H0&
lbl81.ForeColor = &H0&
End Sub
Private Sub Form_Load()
frmInicio.Show
Me.Hide
End Sub
Private Sub lbl10_Click()
lbl10.Caption = cmbNumeros.Text
If lbl10.Caption = lbl11.Caption Or lbl10.Caption = lbl12.Caption Or lbl10.Caption = lbl1.Caption Or lbl10.Caption = lbl2.Caption Or lbl10.Caption = lbl3.Caption Or lbl10.Caption = lbl19.Caption Or lbl10.Caption = lbl20.Caption Or lbl10.Caption = lbl21.Caption Or lbl10.Caption = lbl13.Caption Or lbl10.Caption = lbl16.Caption Or lbl10.Caption = lbl37.Caption Or lbl10.Caption = lbl40.Caption Or lbl10.Caption = lbl43.Caption Or lbl10.Caption = lbl64.Caption Or lbl10.Caption = lbl67.Caption Or lbl10.Caption = lbl70.Caption Or lbl10.Caption = lbl14.Caption Or lbl10.Caption = lbl15.Caption Or lbl10.Caption = lbl17.Caption Or lbl10.Caption = lbl18.Caption Then
    lbl10.ForeColor = &HFF&
Else
    lbl10.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl11_Click()
lbl11.Caption = cmbNumeros.Text
If lbl11.Caption = lbl1.Caption Or lbl11.Caption = lbl2.Caption Or lbl11.Caption = lbl3.Caption Or lbl11.Caption = lbl10.Caption Or lbl11.Caption = lbl12.Caption Or lbl11.Caption = lbl19.Caption Or lbl11.Caption = lbl20.Caption Or lbl11.Caption = lbl21.Caption Or lbl11.Caption = lbl14.Caption Or lbl11.Caption = lbl17.Caption Or lbl11.Caption = lbl38.Caption Or lbl11.Caption = lbl41.Caption Or lbl11.Caption = lbl44.Caption Or lbl11.Caption = lbl65.Caption Or lbl11.Caption = lbl68.Caption Or lbl11.Caption = lbl71.Caption Or lbl11.Caption = lbl13.Caption Or lbl11.Caption = lbl15.Caption Or lbl11.Caption = lbl16.Caption Or lbl11.Caption = lbl18.Caption Then
    lbl11.ForeColor = &HFF&
Else
    lbl11.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl12_Click()
lbl12.Caption = cmbNumeros.Text
If lbl12.Caption = lbl1.Caption Or lbl12.Caption = lbl2.Caption Or lbl12.Caption = lbl3.Caption Or lbl12.Caption = lbl10.Caption Or lbl12.Caption = lbl11.Caption Or lbl12.Caption = lbl19.Caption Or lbl12.Caption = lbl20.Caption Or lbl12.Caption = lbl21.Caption Or lbl12.Caption = lbl15.Caption Or lbl12.Caption = lbl18.Caption Or lbl12.Caption = lbl39.Caption Or lbl12.Caption = lbl42.Caption Or lbl12.Caption = lbl45.Caption Or lbl12.Caption = lbl66.Caption Or lbl12.Caption = lbl69.Caption Or lbl12.Caption = lbl72.Caption Or lbl12.Caption = lbl13.Caption Or lbl12.Caption = lbl14.Caption Or lbl12.Caption = lbl16.Caption Or lbl12.Caption = lbl17.Caption Then
    lbl12.ForeColor = &HFF&
Else
    lbl12.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl13_Click()
lbl13.Caption = cmbNumeros.Text
If lbl13.Caption = lbl4.Caption Or lbl13.Caption = lbl5.Caption Or lbl13.Caption = lbl6.Caption Or lbl13.Caption = lbl14.Caption Or lbl13.Caption = lbl15.Caption Or lbl13.Caption = lbl22.Caption Or lbl13.Caption = lbl23.Caption Or lbl13.Caption = lbl24.Caption Or lbl13.Caption = lbl10.Caption Or lbl13.Caption = lbl16.Caption Or lbl13.Caption = lbl37.Caption Or lbl13.Caption = lbl40.Caption Or lbl13.Caption = lbl43.Caption Or lbl13.Caption = lbl64.Caption Or lbl13.Caption = lbl67.Caption Or lbl13.Caption = lbl70.Caption Or lbl13.Caption = lbl11.Caption Or lbl13.Caption = lbl12.Caption Or lbl13.Caption = lbl17.Caption Or lbl13.Caption = lbl18.Caption Then
    lbl13.ForeColor = &HFF&
Else
    lbl13.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl15_Click()
lbl15.Caption = cmbNumeros.Text
If lbl15.Caption = lbl10.Caption Or lbl15.Caption = lbl11.Caption Or lbl15.Caption = lbl12.Caption Or lbl15.Caption = lbl13.Caption Or lbl15.Caption = lbl14.Caption Or lbl15.Caption = lbl16.Caption Or lbl15.Caption = lbl17.Caption Or lbl15.Caption = lbl18.Caption Or lbl15.Caption = lbl4.Caption Or lbl15.Caption = lbl5.Caption Or lbl15.Caption = lbl6.Caption Or lbl15.Caption = lbl22.Caption Or lbl15.Caption = lbl23.Caption Or lbl15.Caption = lbl24.Caption Or lbl15.Caption = lbl39.Caption Or lbl15.Caption = lbl42.Caption Or lbl15.Caption = lbl45.Caption Or lbl15.Caption = lbl66.Caption Or lbl15.Caption = lbl69.Caption Or lbl15.Caption = lbl72.Caption Then
    lbl15.ForeColor = &HFF&
Else
    lbl15.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl16_Click()
lbl16.Caption = cmbNumeros.Text
If lbl16.Caption = lbl10.Caption Or lbl16.Caption = lbl11.Caption Or lbl16.Caption = lbl12.Caption Or lbl16.Caption = lbl13.Caption Or lbl16.Caption = lbl14.Caption Or lbl16.Caption = lbl15.Caption Or lbl16.Caption = lbl17.Caption Or lbl16.Caption = lbl18.Caption Or lbl16.Caption = lbl7.Caption Or lbl16.Caption = lbl8.Caption Or lbl16.Caption = lbl9.Caption Or lbl16.Caption = lbl25.Caption Or lbl16.Caption = lbl26.Caption Or lbl16.Caption = lbl27.Caption Or lbl16.Caption = lbl37.Caption Or lbl16.Caption = lbl40.Caption Or lbl16.Caption = lbl43.Caption Or lbl16.Caption = lbl64.Caption Or lbl16.Caption = lbl67.Caption Or lbl16.Caption = lbl70.Caption Then
    lbl16.ForeColor = &HFF&
Else
    lbl16.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl17_Click()
lbl17.Caption = cmbNumeros.Text
If lbl17.Caption = lbl7.Caption Or lbl17.Caption = lbl8.Caption Or lbl17.Caption = lbl9.Caption Or lbl17.Caption = lbl25.Caption Or lbl17.Caption = lbl26.Caption Or lbl17.Caption = lbl27.Caption Or lbl17.Caption = lbl10.Caption Or lbl17.Caption = lbl11.Caption Or lbl17.Caption = lbl12.Caption Or lbl17.Caption = lbl13.Caption Or lbl17.Caption = lbl14.Caption Or lbl17.Caption = lbl15.Caption Or lbl17.Caption = lbl16.Caption Or lbl17.Caption = lbl18.Caption Or lbl17.Caption = lbl38.Caption Or lbl17.Caption = lbl41.Caption Or lbl17.Caption = lbl44.Caption Or lbl17.Caption = lbl65.Caption Or lbl17.Caption = lbl68.Caption Or lbl17.Caption = lbl71.Caption Then
    lbl17.ForeColor = &HFF&
Else
    lbl17.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl18_Click()
lbl18.Caption = cmbNumeros.Text
If lbl18.Caption = lbl7.Caption Or lbl18.Caption = lbl8.Caption Or lbl18.Caption = lbl9.Caption Or lbl18.Caption = lbl25.Caption Or lbl18.Caption = lbl26.Caption Or lbl18.Caption = lbl27.Caption Or lbl18.Caption = lbl10.Caption Or lbl18.Caption = lbl11.Caption Or lbl18.Caption = lbl12.Caption Or lbl18.Caption = lbl13.Caption Or lbl18.Caption = lbl14.Caption Or lbl18.Caption = lbl15.Caption Or lbl18.Caption = lbl16.Caption Or lbl18.Caption = lbl17.Caption Or lbl18.Caption = lbl39.Caption Or lbl18.Caption = lbl42.Caption Or lbl18.Caption = lbl45.Caption Or lbl18.Caption = lbl66.Caption Or lbl18.Caption = lbl69.Caption Or lbl18.Caption = lbl72.Caption Then
    lbl18.ForeColor = &HFF&
Else
    lbl18.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl19_Click()
lbl19.Caption = cmbNumeros.Text
If lbl19.Caption = lbl1.Caption Or lbl19.Caption = lbl2.Caption Or lbl19.Caption = lbl3.Caption Or lbl19.Caption = lbl10.Caption Or lbl19.Caption = lbl11.Caption Or lbl19.Caption = lbl12.Caption Or lbl19.Caption = lbl20.Caption Or lbl19.Caption = lbl21.Caption Or lbl19.Caption = lbl22.Caption Or lbl19.Caption = lbl23.Caption Or lbl19.Caption = lbl24.Caption Or lbl19.Caption = lbl25.Caption Or lbl19.Caption = lbl26.Caption Or lbl19.Caption = lbl27.Caption Or lbl19.Caption = lbl46.Caption Or lbl19.Caption = lbl49.Caption Or lbl19.Caption = lbl52.Caption Or lbl19.Caption = lbl73.Caption Or lbl19.Caption = lbl76.Caption Or lbl19.Caption = lbl79.Caption Then
    lbl19.ForeColor = &HFF&
Else
    lbl18.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl2_Click()
lbl2.Caption = cmbNumeros.Text
If lbl2.Caption = lbl10.Caption Or lbl2.Caption = lbl11.Caption Or lbl2.Caption = lbl12.Caption Or lbl2.Caption = lbl19.Caption Or lbl2.Caption = lbl20.Caption Or lbl2.Caption = lbl21.Caption Or lbl2.Caption = lbl1.Caption Or lbl2.Caption = lbl3.Caption Or lbl2.Caption = lbl4.Caption Or lbl2.Caption = lbl5.Caption Or lbl2.Caption = lbl6.Caption Or lbl2.Caption = lbl7.Caption Or lbl2.Caption = lbl8.Caption Or lbl2.Caption = lbl9.Caption Or lbl2.Caption = lbl29.Caption Or lbl2.Caption = lbl32.Caption Or lbl2.Caption = lbl35.Caption Or lbl2.Caption = lbl56.Caption Or lbl2.Caption = lbl59.Caption Or lbl2.Caption = lbl62.Caption Then
    lbl2.ForeColor = &HFF&
Else
    lbl2.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl20_Click()
lbl20.Caption = cmbNumeros.Text
If lbl20.Caption = lbl1.Caption Or lbl20.Caption = lbl2.Caption Or lbl20.Caption = lbl3.Caption Or lbl20.Caption = lbl10.Caption Or lbl20.Caption = lbl11.Caption Or lbl20.Caption = lbl12.Caption Or lbl20.Caption = lbl19.Caption Or lbl20.Caption = lbl21.Caption Or lbl20.Caption = lbl22.Caption Or lbl20.Caption = lbl23.Caption Or lbl20.Caption = lbl24.Caption Or lbl20.Caption = lbl25.Caption Or lbl20.Caption = lbl26.Caption Or lbl20.Caption = lbl27.Caption Or lbl20.Caption = lbl47.Caption Or lbl20.Caption = lbl50.Caption Or lbl20.Caption = lbl53.Caption Or lbl20.Caption = lbl74.Caption Or lbl20.Caption = lbl77.Caption Or lbl20.Caption = lbl80.Caption Then
    lbl20.ForeColor = &HFF&
Else
    lbl20.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl21_Click()
lbl21.Caption = cmbNumeros.Text
If lbl21.Caption = lbl1.Caption Or lbl21.Caption = lbl2.Caption Or lbl21.Caption = lbl3.Caption Or lbl21.Caption = lbl10.Caption Or lbl21.Caption = lbl11.Caption Or lbl21.Caption = lbl12.Caption Or lbl21.Caption = lbl19.Caption Or lbl21.Caption = lbl20.Caption Or lbl21.Caption = lbl22.Caption Or lbl21.Caption = lbl23.Caption Or lbl21.Caption = lbl24.Caption Or lbl21.Caption = lbl25.Caption Or lbl21.Caption = lbl26.Caption Or lbl21.Caption = lbl27.Caption Or lbl21.Caption = lbl48.Caption Or lbl21.Caption = lbl51.Caption Or lbl21.Caption = lbl54.Caption Or lbl21.Caption = lbl75.Caption Or lbl21.Caption = lbl78.Caption Or lbl21.Caption = lbl81.Caption Then
    lbl21.ForeColor = &HFF&
Else
    lbl21.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl22_Click()
lbl22.Caption = cmbNumeros.Text
If lbl22.Caption = lbl4.Caption Or lbl22.Caption = lbl5.Caption Or lbl22.Caption = lbl6.Caption Or lbl22.Caption = lbl13.Caption Or lbl22.Caption = lbl14.Caption Or lbl22.Caption = lbl15.Caption Or lbl22.Caption = lbl19.Caption Or lbl22.Caption = lbl20.Caption Or lbl22.Caption = lbl21.Caption Or lbl22.Caption = lbl23.Caption Or lbl22.Caption = lbl24.Caption Or lbl22.Caption = lbl25.Caption Or lbl22.Caption = lbl26.Caption Or lbl22.Caption = lbl27.Caption Or lbl22.Caption = lbl46.Caption Or lbl22.Caption = lbl49.Caption Or lbl22.Caption = lbl52.Caption Or lbl22.Caption = lbl73.Caption Or lbl22.Caption = lbl76.Caption Or lbl22.Caption = lbl79.Caption Then
    lbl22.ForeColor = &HFF&
Else
    lbl22.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl23_Click()
lbl23.Caption = cmbNumeros.Text
If lbl23.Caption = lbl4.Caption Or lbl23.Caption = lbl5.Caption Or lbl23.Caption = lbl6.Caption Or lbl23.Caption = lbl13.Caption Or lbl23.Caption = lbl14.Caption Or lbl23.Caption = lbl15.Caption Or lbl23.Caption = lbl19.Caption Or lbl23.Caption = lbl20.Caption Or lbl23.Caption = lbl21.Caption Or lbl23.Caption = lbl22.Caption Or lbl23.Caption = lbl24.Caption Or lbl23.Caption = lbl25.Caption Or lbl23.Caption = lbl26.Caption Or lbl23.Caption = lbl27.Caption Or lbl23.Caption = lbl47.Caption Or lbl23.Caption = lbl50.Caption Or lbl23.Caption = lbl53.Caption Or lbl23.Caption = lbl74.Caption Or lbl23.Caption = lbl77.Caption Or lbl23.Caption = lbl80.Caption Then
    lbl23.ForeColor = &HFF&
Else
    lbl23.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl24_Click()
lbl24.Caption = cmbNumeros.Text
If lbl24.Caption = lbl4.Caption Or lbl24.Caption = lbl5.Caption Or lbl24.Caption = lbl6.Caption Or lbl24.Caption = lbl13.Caption Or lbl24.Caption = lbl14.Caption Or lbl24.Caption = lbl15.Caption Or lbl24.Caption = lbl19.Caption Or lbl24.Caption = lbl20.Caption Or lbl24.Caption = lbl21.Caption Or lbl24.Caption = lbl22.Caption Or lbl24.Caption = lbl23.Caption Or lbl24.Caption = lbl25.Caption Or lbl24.Caption = lbl26.Caption Or lbl24.Caption = lbl27.Caption Or lbl24.Caption = lbl48.Caption Or lbl24.Caption = lbl51.Caption Or lbl24.Caption = lbl54.Caption Or lbl24.Caption = lbl75.Caption Or lbl24.Caption = lbl78.Caption Or lbl24.Caption = lbl81.Caption Then
    lbl24.ForeColor = &HFF&
Else
    lbl24.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl25_Click()
lbl25.Caption = cmbNumeros.Text
If lbl25.Caption = lbl7.Caption Or lbl25.Caption = lbl8.Caption Or lbl25.Caption = lbl9.Caption Or lbl25.Caption = lbl16.Caption Or lbl25.Caption = lbl17.Caption Or lbl25.Caption = lbl18.Caption Or lbl25.Caption = lbl19.Caption Or lbl25.Caption = lbl20.Caption Or lbl25.Caption = lbl21.Caption Or lbl25.Caption = lbl22.Caption Or lbl25.Caption = lbl23.Caption Or lbl25.Caption = lbl24.Caption Or lbl25.Caption = lbl26.Caption Or lbl25.Caption = lbl27.Caption Or lbl25.Caption = lbl46.Caption Or lbl25.Caption = lbl49.Caption Or lbl25.Caption = lbl52.Caption Or lbl25.Caption = lbl73.Caption Or lbl25.Caption = lbl76.Caption Or lbl25.Caption = lbl79.Caption Then
    lbl25.ForeColor = &HFF&
Else
    lbl25.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl26_Click()
lbl26.Caption = cmbNumeros.Text
If lbl26.Caption = lbl7.Caption Or lbl26.Caption = lbl8.Caption Or lbl26.Caption = lbl9.Caption Or lbl26.Caption = lbl16.Caption Or lbl26.Caption = lbl17.Caption Or lbl26.Caption = lbl18.Caption Or lbl26.Caption = lbl19.Caption Or lbl26.Caption = lbl20.Caption Or lbl26.Caption = lbl21.Caption Or lbl26.Caption = lbl22.Caption Or lbl26.Caption = lbl23.Caption Or lbl26.Caption = lbl24.Caption Or lbl26.Caption = lbl25.Caption Or lbl26.Caption = lbl27.Caption Or lbl26.Caption = lbl47.Caption Or lbl26.Caption = lbl50.Caption Or lbl26.Caption = lbl53.Caption Or lbl26.Caption = lbl74.Caption Or lbl26.Caption = lbl77.Caption Or lbl26.Caption = lbl80.Caption Then
    lbl26.ForeColor = &HFF&
Else
    lbl26.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl28_Click()
lbl28.Caption = cmbNumeros.Text
If lbl28.Caption = lbl1.Caption Or lbl28.Caption = lbl4.Caption Or lbl28.Caption = lbl7.Caption Or lbl28.Caption = lbl55.Caption Or lbl28.Caption = lbl58.Caption Or lbl28.Caption = lbl61.Caption Or lbl28.Caption = lbl29.Caption Or lbl28.Caption = lbl30.Caption Or lbl28.Caption = lbl31.Caption Or lbl28.Caption = lbl32.Caption Or lbl28.Caption = lbl33.Caption Or lbl28.Caption = lbl34.Caption Or lbl28.Caption = lbl35.Caption Or lbl28.Caption = lbl36.Caption Or lbl28.Caption = lbl37.Caption Or lbl28.Caption = lbl38.Caption Or lbl28.Caption = lbl39.Caption Or lbl28.Caption = lbl46.Caption Or lbl28.Caption = lbl47.Caption Or lbl28.Caption = lbl48.Caption Then
    lbl28.ForeColor = &HFF&
Else
    lbl28.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl29_Click()
lbl29.Caption = cmbNumeros.Text
If lbl29.Caption = lbl2.Caption Or lbl29.Caption = lbl5.Caption Or lbl29.Caption = lbl8.Caption Or lbl29.Caption = lbl56.Caption Or lbl29.Caption = lbl59.Caption Or lbl29.Caption = lbl62.Caption Or lbl29.Caption = lbl28.Caption Or lbl29.Caption = lbl30.Caption Or lbl29.Caption = lbl31.Caption Or lbl29.Caption = lbl32.Caption Or lbl29.Caption = lbl33.Caption Or lbl29.Caption = lbl34.Caption Or lbl29.Caption = lbl35.Caption Or lbl29.Caption = lbl36.Caption Or lbl29.Caption = lbl37.Caption Or lbl29.Caption = lbl38.Caption Or lbl29.Caption = lbl39.Caption Or lbl29.Caption = lbl46.Caption Or lbl29.Caption = lbl47.Caption Or lbl29.Caption = lbl48.Caption Then
    lbl29.ForeColor = &HFF&
Else
    lbl29.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl3_Click()
lbl3.Caption = cmbNumeros.Text
If lbl3.Caption = lbl10.Caption Or lbl3.Caption = lbl11.Caption Or lbl3.Caption = lbl12.Caption Or lbl3.Caption = lbl19.Caption Or lbl3.Caption = lbl20.Caption Or lbl3.Caption = lbl21.Caption Or lbl3.Caption = lbl1.Caption Or lbl3.Caption = lbl2.Caption Or lbl3.Caption = lbl4.Caption Or lbl3.Caption = lbl5.Caption Or lbl3.Caption = lbl6.Caption Or lbl3.Caption = lbl7.Caption Or lbl3.Caption = lbl8.Caption Or lbl3.Caption = lbl9.Caption Or lbl3.Caption = lbl30.Caption Or lbl3.Caption = lbl33.Caption Or lbl3.Caption = lbl36.Caption Or lbl3.Caption = lbl57.Caption Or lbl3.Caption = lbl60.Caption Or lbl3.Caption = lbl63.Caption Then
    lbl3.ForeColor = &HFF&
Else
    lbl3.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl30_Click()
lbl30.Caption = cmbNumeros.Text
If lbl30.Caption = lbl3.Caption Or lbl30.Caption = lbl6.Caption Or lbl30.Caption = lbl9.Caption Or lbl30.Caption = lbl57.Caption Or lbl30.Caption = lbl60.Caption Or lbl30.Caption = lbl63.Caption Or lbl30.Caption = lbl28.Caption Or lbl30.Caption = lbl29.Caption Or lbl30.Caption = lbl31.Caption Or lbl30.Caption = lbl32.Caption Or lbl30.Caption = lbl33.Caption Or lbl30.Caption = lbl34.Caption Or lbl30.Caption = lbl35.Caption Or lbl30.Caption = lbl36.Caption Or lbl30.Caption = lbl37.Caption Or lbl30.Caption = lbl38.Caption Or lbl30.Caption = lbl39.Caption Or lbl30.Caption = lbl46.Caption Or lbl30.Caption = lbl47.Caption Or lbl30.Caption = lbl48.Caption Then
    lbl30.ForeColor = &HFF&
Else
    lbl30.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl31_Click()
lbl31.Caption = cmbNumeros.Text
If lbl31.Caption = lbl1.Caption Or lbl31.Caption = lbl4.Caption Or lbl31.Caption = lbl7.Caption Or lbl31.Caption = lbl55.Caption Or lbl31.Caption = lbl58.Caption Or lbl31.Caption = lbl61.Caption Or lbl31.Caption = lbl28.Caption Or lbl31.Caption = lbl29.Caption Or lbl31.Caption = lbl30.Caption Or lbl31.Caption = lbl32.Caption Or lbl31.Caption = lbl33.Caption Or lbl31.Caption = lbl34.Caption Or lbl31.Caption = lbl35.Caption Or lbl31.Caption = lbl36.Caption Or lbl31.Caption = lbl40.Caption Or lbl31.Caption = lbl41.Caption Or lbl31.Caption = lbl42.Caption Or lbl31.Caption = lbl49.Caption Or lbl31.Caption = lbl50.Caption Or lbl31.Caption = lbl51.Caption Then
    lbl31.ForeColor = &HFF&
Else
    lbl31.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl33_Click()
lbl33.Caption = cmbNumeros.Text
If lbl33.Caption = lbl3.Caption Or lbl33.Caption = lbl6.Caption Or lbl33.Caption = lbl9.Caption Or lbl33.Caption = lbl57.Caption Or lbl33.Caption = lbl60.Caption Or lbl33.Caption = lbl63.Caption Or lbl33.Caption = lbl28.Caption Or lbl33.Caption = lbl29.Caption Or lbl33.Caption = lbl30.Caption Or lbl33.Caption = lbl31.Caption Or lbl33.Caption = lbl32.Caption Or lbl33.Caption = lbl34.Caption Or lbl33.Caption = lbl35.Caption Or lbl33.Caption = lbl36.Caption Or lbl33.Caption = lbl40.Caption Or lbl33.Caption = lbl41.Caption Or lbl33.Caption = lbl42.Caption Or lbl33.Caption = lbl49.Caption Or lbl33.Caption = lbl50.Caption Or lbl33.Caption = lbl51.Caption Then
    lbl33.ForeColor = &HFF&
Else
    lbl33.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl34_Click()
lbl34.Caption = cmbNumeros.Text
If lbl34.Caption = lbl1.Caption Or lbl34.Caption = lbl4.Caption Or lbl34.Caption = lbl7.Caption Or lbl34.Caption = lbl55.Caption Or lbl34.Caption = lbl58.Caption Or lbl34.Caption = lbl61.Caption Or lbl34.Caption = lbl28.Caption Or lbl34.Caption = lbl29.Caption Or lbl34.Caption = lbl30.Caption Or lbl34.Caption = lbl31.Caption Or lbl34.Caption = lbl32.Caption Or lbl34.Caption = lbl33.Caption Or lbl34.Caption = lbl35.Caption Or lbl34.Caption = lbl36.Caption Or lbl34.Caption = lbl43.Caption Or lbl34.Caption = lbl44.Caption Or lbl34.Caption = lbl45.Caption Or lbl34.Caption = lbl52.Caption Or lbl34.Caption = lbl53.Caption Or lbl34.Caption = lbl54.Caption Then
    lbl34.ForeColor = &HFF&
Else
    lbl34.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl35_Click()
lbl35.Caption = cmbNumeros.Text
If lbl35.Caption = lbl2.Caption Or lbl35.Caption = lbl5.Caption Or lbl35.Caption = lbl8.Caption Or lbl35.Caption = lbl56.Caption Or lbl35.Caption = lbl59.Caption Or lbl35.Caption = lbl62.Caption Or lbl35.Caption = lbl28.Caption Or lbl35.Caption = lbl29.Caption Or lbl35.Caption = lbl30.Caption Or lbl35.Caption = lbl31.Caption Or lbl35.Caption = lbl32.Caption Or lbl35.Caption = lbl33.Caption Or lbl35.Caption = lbl34.Caption Or lbl35.Caption = lbl36.Caption Or lbl35.Caption = lbl43.Caption Or lbl35.Caption = lbl44.Caption Or lbl35.Caption = lbl45.Caption Or lbl35.Caption = lbl52.Caption Or lbl35.Caption = lbl53.Caption Or lbl35.Caption = lbl54.Caption Then
    lbl35.ForeColor = &HFF&
Else
    lbl35.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl36_Click()
lbl36.Caption = cmbNumeros.Text
If lbl36.Caption = lbl3.Caption Or lbl36.Caption = lbl6.Caption Or lbl36.Caption = lbl9.Caption Or lbl36.Caption = lbl57.Caption Or lbl36.Caption = lbl60.Caption Or lbl36.Caption = lbl63.Caption Or lbl36.Caption = lbl28.Caption Or lbl36.Caption = lbl29.Caption Or lbl36.Caption = lbl30.Caption Or lbl36.Caption = lbl31.Caption Or lbl36.Caption = lbl32.Caption Or lbl36.Caption = lbl33.Caption Or lbl36.Caption = lbl34.Caption Or lbl36.Caption = lbl35.Caption Or lbl36.Caption = lbl43.Caption Or lbl36.Caption = lbl44.Caption Or lbl36.Caption = lbl45.Caption Or lbl36.Caption = lbl52.Caption Or lbl36.Caption = lbl53.Caption Or lbl36.Caption = lbl54.Caption Then
    lbl36.ForeColor = &HFF&
Else
    lbl36.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl37_Click()
lbl37.Caption = cmbNumeros.Text
If lbl37.Caption = lbl10.Caption Or lbl37.Caption = lbl13.Caption Or lbl37.Caption = lbl16.Caption Or lbl37.Caption = lbl64.Caption Or lbl37.Caption = lbl67.Caption Or lbl37.Caption = lbl70.Caption Or lbl37.Caption = lbl38.Caption Or lbl37.Caption = lbl39.Caption Or lbl37.Caption = lbl40.Caption Or lbl37.Caption = lbl41.Caption Or lbl37.Caption = lbl42.Caption Or lbl37.Caption = lbl43.Caption Or lbl37.Caption = lbl44.Caption Or lbl37.Caption = lbl45.Caption Or lbl37.Caption = lbl28.Caption Or lbl37.Caption = lbl29.Caption Or lbl37.Caption = lbl30.Caption Or lbl37.Caption = lbl46.Caption Or lbl37.Caption = lbl47.Caption Or lbl37.Caption = lbl48.Caption Then
    lbl37.ForeColor = &HFF&
Else
    lbl37.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl39_Click()
lbl39.Caption = cmbNumeros.Text
If lbl39.Caption = lbl12.Caption Or lbl39.Caption = lbl15.Caption Or lbl39.Caption = lbl18.Caption Or lbl39.Caption = lbl66.Caption Or lbl39.Caption = lbl69.Caption Or lbl39.Caption = lbl72.Caption Or lbl39.Caption = lbl37.Caption Or lbl39.Caption = lbl38.Caption Or lbl39.Caption = lbl40.Caption Or lbl39.Caption = lbl41.Caption Or lbl39.Caption = lbl42.Caption Or lbl39.Caption = lbl43.Caption Or lbl39.Caption = lbl44.Caption Or lbl39.Caption = lbl45.Caption Or lbl39.Caption = lbl28.Caption Or lbl39.Caption = lbl29.Caption Or lbl39.Caption = lbl30.Caption Or lbl39.Caption = lbl46.Caption Or lbl39.Caption = lbl47.Caption Or lbl39.Caption = lbl48.Caption Then
    lbl39.ForeColor = &HFF&
Else
    lbl39.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl4_Click()
lbl4.Caption = cmbNumeros.Text
If lbl4.Caption = lbl13.Caption Or lbl4.Caption = lbl14.Caption Or lbl4.Caption = lbl15.Caption Or lbl4.Caption = lbl22.Caption Or lbl4.Caption = lbl23.Caption Or lbl4.Caption = lbl24.Caption Or lbl4.Caption = lbl1.Caption Or lbl4.Caption = lbl2.Caption Or lbl4.Caption = lbl3.Caption Or lbl4.Caption = lbl5.Caption Or lbl4.Caption = lbl6.Caption Or lbl4.Caption = lbl7.Caption Or lbl4.Caption = lbl8.Caption Or lbl4.Caption = lbl9.Caption Or lbl4.Caption = lbl28.Caption Or lbl4.Caption = lbl31.Caption Or lbl4.Caption = lbl34.Caption Or lbl4.Caption = lbl55.Caption Or lbl4.Caption = lbl58.Caption Or lbl4.Caption = lbl61.Caption Then
    lbl4.ForeColor = &HFF&
Else
    lbl4.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl40_Click()
lbl40.Caption = cmbNumeros.Text
If lbl40.Caption = lbl10.Caption Or lbl40.Caption = lbl13.Caption Or lbl40.Caption = lbl16.Caption Or lbl40.Caption = lbl64.Caption Or lbl40.Caption = lbl67.Caption Or lbl40.Caption = lbl70.Caption Or lbl40.Caption = lbl37.Caption Or lbl40.Caption = lbl38.Caption Or lbl40.Caption = lbl39.Caption Or lbl40.Caption = lbl41.Caption Or lbl40.Caption = lbl42.Caption Or lbl40.Caption = lbl43.Caption Or lbl40.Caption = lbl44.Caption Or lbl40.Caption = lbl45.Caption Or lbl40.Caption = lbl31.Caption Or lbl40.Caption = lbl32.Caption Or lbl40.Caption = lbl33.Caption Or lbl40.Caption = lbl49.Caption Or lbl40.Caption = lbl50.Caption Or lbl40.Caption = lbl51.Caption Then
    lbl40.ForeColor = &HFF&
Else
    lbl40.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl41_Click()
lbl41.Caption = cmbNumeros.Text
If lbl41.Caption = lbl11.Caption Or lbl41.Caption = lbl14.Caption Or lbl41.Caption = lbl17.Caption Or lbl41.Caption = lbl65.Caption Or lbl41.Caption = lbl68.Caption Or lbl41.Caption = lbl71.Caption Or lbl41.Caption = lbl37.Caption Or lbl41.Caption = lbl38.Caption Or lbl41.Caption = lbl39.Caption Or lbl41.Caption = lbl40.Caption Or lbl41.Caption = lbl42.Caption Or lbl41.Caption = lbl43.Caption Or lbl41.Caption = lbl44.Caption Or lbl41.Caption = lbl45.Caption Or lbl41.Caption = lbl31.Caption Or lbl41.Caption = lbl32.Caption Or lbl41.Caption = lbl33.Caption Or lbl41.Caption = lbl49.Caption Or lbl41.Caption = lbl50.Caption Or lbl41.Caption = lbl51.Caption Then
    lbl41.ForeColor = &HFF&
Else
    lbl41.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl42_Click()
lbl42.Caption = cmbNumeros.Text
If lbl42.Caption = lbl12.Caption Or lbl42.Caption = lbl15.Caption Or lbl42.Caption = lbl18.Caption Or lbl42.Caption = lbl66.Caption Or lbl42.Caption = lbl69.Caption Or lbl42.Caption = lbl72.Caption Or lbl42.Caption = lbl37.Caption Or lbl42.Caption = lbl38.Caption Or lbl42.Caption = lbl39.Caption Or lbl42.Caption = lbl40.Caption Or lbl42.Caption = lbl41.Caption Or lbl42.Caption = lbl43.Caption Or lbl42.Caption = lbl44.Caption Or lbl42.Caption = lbl45.Caption Or lbl42.Caption = lbl31.Caption Or lbl42.Caption = lbl32.Caption Or lbl42.Caption = lbl33.Caption Or lbl42.Caption = lbl49.Caption Or lbl42.Caption = lbl50.Caption Or lbl42.Caption = lbl51.Caption Then
    lbl42.ForeColor = &HFF&
Else
    lbl42.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl43_Click()
lbl43.Caption = cmbNumeros.Text
If lbl43.Caption = lbl10.Caption Or lbl43.Caption = lbl13.Caption Or lbl43.Caption = lbl16.Caption Or lbl43.Caption = lbl64.Caption Or lbl43.Caption = lbl67.Caption Or lbl43.Caption = lbl70.Caption Or lbl43.Caption = lbl37.Caption Or lbl43.Caption = lbl38.Caption Or lbl43.Caption = lbl39.Caption Or lbl43.Caption = lbl40.Caption Or lbl43.Caption = lbl41.Caption Or lbl43.Caption = lbl42.Caption Or lbl43.Caption = lbl44.Caption Or lbl43.Caption = lbl45.Caption Or lbl43.Caption = lbl34.Caption Or lbl43.Caption = lbl35.Caption Or lbl43.Caption = lbl36.Caption Or lbl43.Caption = lbl52.Caption Or lbl43.Caption = lbl53.Caption Or lbl43.Caption = lbl54.Caption Then
    lbl43.ForeColor = &HFF&
Else
    lbl43.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl44_Click()
lbl44.Caption = cmbNumeros.Text
If lbl44.Caption = lbl11.Caption Or lbl44.Caption = lbl14.Caption Or lbl44.Caption = lbl17.Caption Or lbl44.Caption = lbl65.Caption Or lbl44.Caption = lbl68.Caption Or lbl44.Caption = lbl71.Caption Or lbl44.Caption = lbl37.Caption Or lbl44.Caption = lbl38.Caption Or lbl44.Caption = lbl39.Caption Or lbl44.Caption = lbl40.Caption Or lbl44.Caption = lbl41.Caption Or lbl44.Caption = lbl42.Caption Or lbl44.Caption = lbl43.Caption Or lbl44.Caption = lbl45.Caption Or lbl44.Caption = lbl34.Caption Or lbl44.Caption = lbl35.Caption Or lbl44.Caption = lbl36.Caption Or lbl44.Caption = lbl52.Caption Or lbl44.Caption = lbl53.Caption Or lbl44.Caption = lbl54.Caption Then
    lbl44.ForeColor = &HFF&
Else
    lbl44.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl45_Click()
lbl45.Caption = cmbNumeros.Text
If lbl45.Caption = lbl12.Caption Or lbl45.Caption = lbl15.Caption Or lbl45.Caption = lbl18.Caption Or lbl45.Caption = lbl66.Caption Or lbl45.Caption = lbl69.Caption Or lbl45.Caption = lbl72.Caption Or lbl45.Caption = lbl37.Caption Or lbl45.Caption = lbl38.Caption Or lbl45.Caption = lbl39.Caption Or lbl45.Caption = lbl40.Caption Or lbl45.Caption = lbl41.Caption Or lbl45.Caption = lbl42.Caption Or lbl45.Caption = lbl43.Caption Or lbl45.Caption = lbl44.Caption Or lbl45.Caption = lbl34.Caption Or lbl45.Caption = lbl35.Caption Or lbl45.Caption = lbl36.Caption Or lbl45.Caption = lbl52.Caption Or lbl45.Caption = lbl53.Caption Or lbl45.Caption = lbl54.Caption Then
    lbl45.ForeColor = &HFF&
Else
    lbl45.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl46_Click()
lbl46.Caption = cmbNumeros.Text
If lbl46.Caption = lbl19.Caption Or lbl46.Caption = lbl22.Caption Or lbl46.Caption = lbl25.Caption Or lbl46.Caption = lbl73.Caption Or lbl46.Caption = lbl76.Caption Or lbl46.Caption = lbl79.Caption Or lbl46.Caption = lbl47.Caption Or lbl46.Caption = lbl48.Caption Or lbl46.Caption = lbl49.Caption Or lbl46.Caption = lbl50.Caption Or lbl46.Caption = lbl51.Caption Or lbl46.Caption = lbl52.Caption Or lbl46.Caption = lbl53.Caption Or lbl46.Caption = lbl54.Caption Or lbl46.Caption = lbl28.Caption Or lbl46.Caption = lbl29.Caption Or lbl46.Caption = lbl30.Caption Or lbl46.Caption = lbl37.Caption Or lbl46.Caption = lbl38.Caption Or lbl46.Caption = lbl39.Caption Then
    lbl46.ForeColor = &HFF&
Else
    lbl46.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl47_Click()
lbl47.Caption = cmbNumeros.Text
If lbl47.Caption = lbl20.Caption Or lbl47.Caption = lbl23.Caption Or lbl47.Caption = lbl26.Caption Or lbl47.Caption = lbl74.Caption Or lbl47.Caption = lbl77.Caption Or lbl47.Caption = lbl80.Caption Or lbl47.Caption = lbl46.Caption Or lbl47.Caption = lbl48.Caption Or lbl47.Caption = lbl49.Caption Or lbl47.Caption = lbl50.Caption Or lbl47.Caption = lbl51.Caption Or lbl47.Caption = lbl52.Caption Or lbl47.Caption = lbl53.Caption Or lbl47.Caption = lbl54.Caption Or lbl47.Caption = lbl28.Caption Or lbl47.Caption = lbl29.Caption Or lbl47.Caption = lbl30.Caption Or lbl47.Caption = lbl37.Caption Or lbl47.Caption = lbl38.Caption Or lbl47.Caption = lbl39.Caption Then
    lbl47.ForeColor = &HFF&
Else
    lbl47.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl48_Click()
lbl48.Caption = cmbNumeros.Text
If lbl48.Caption = lbl21.Caption Or lbl48.Caption = lbl24.Caption Or lbl48.Caption = lbl27.Caption Or lbl48.Caption = lbl75.Caption Or lbl48.Caption = lbl78.Caption Or lbl48.Caption = lbl81.Caption Or lbl48.Caption = lbl46.Caption Or lbl48.Caption = lbl47.Caption Or lbl48.Caption = lbl49.Caption Or lbl48.Caption = lbl50.Caption Or lbl48.Caption = lbl51.Caption Or lbl48.Caption = lbl52.Caption Or lbl48.Caption = lbl53.Caption Or lbl48.Caption = lbl54.Caption Or lbl48.Caption = lbl28.Caption Or lbl48.Caption = lbl29.Caption Or lbl48.Caption = lbl30.Caption Or lbl48.Caption = lbl37.Caption Or lbl48.Caption = lbl38.Caption Or lbl48.Caption = lbl39.Caption Then
    lbl48.ForeColor = &HFF&
Else
    lbl48.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl49_Click()
lbl49.Caption = cmbNumeros.Text
If lbl49.Caption = lbl19.Caption Or lbl49.Caption = lbl22.Caption Or lbl49.Caption = lbl25.Caption Or lbl49.Caption = lbl73.Caption Or lbl49.Caption = lbl76.Caption Or lbl49.Caption = lbl79.Caption Or lbl49.Caption = lbl46.Caption Or lbl49.Caption = lbl47.Caption Or lbl49.Caption = lbl48.Caption Or lbl49.Caption = lbl50.Caption Or lbl49.Caption = lbl51.Caption Or lbl49.Caption = lbl52.Caption Or lbl49.Caption = lbl53.Caption Or lbl49.Caption = lbl54.Caption Or lbl49.Caption = lbl31.Caption Or lbl49.Caption = lbl32.Caption Or lbl49.Caption = lbl33.Caption Or lbl49.Caption = lbl40.Caption Or lbl49.Caption = lbl41.Caption Or lbl49.Caption = lbl42.Caption Then
    lbl49.ForeColor = &HFF&
Else
    lbl49.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl5_Click()
lbl5.Caption = cmbNumeros.Text
If lbl5.Caption = lbl29.Caption Or lbl5.Caption = lbl32.Caption Or lbl5.Caption = lbl35.Caption Or lbl5.Caption = lbl56.Caption Or lbl5.Caption = lbl59.Caption Or lbl5.Caption = lbl62.Caption Or lbl5.Caption = lbl1.Caption Or lbl5.Caption = lbl2.Caption Or lbl5.Caption = lbl3.Caption Or lbl5.Caption = lbl4.Caption Or lbl5.Caption = lbl6.Caption Or lbl5.Caption = lbl7.Caption Or lbl5.Caption = lbl8.Caption Or lbl5.Caption = lbl9.Caption Or lbl5.Caption = lbl13.Caption Or lbl5.Caption = lbl14.Caption Or lbl5.Caption = lbl1.Caption Or lbl5.Caption = lbl22.Caption Or lbl5.Caption = lbl23.Caption Or lbl5.Caption = lbl24.Caption Then
    lbl5.ForeColor = &HFF&
Else
    lbl5.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl50_Click()
lbl50.Caption = cmbNumeros.Text
If lbl50.Caption = lbl20.Caption Or lbl50.Caption = lbl23.Caption Or lbl50.Caption = lbl26.Caption Or lbl50.Caption = lbl74.Caption Or lbl50.Caption = lbl77.Caption Or lbl50.Caption = lbl80.Caption Or lbl50.Caption = lbl46.Caption Or lbl50.Caption = lbl47.Caption Or lbl50.Caption = lbl48.Caption Or lbl50.Caption = lbl49.Caption Or lbl50.Caption = lbl51.Caption Or lbl50.Caption = lbl52.Caption Or lbl50.Caption = lbl53.Caption Or lbl50.Caption = lbl54.Caption Or lbl50.Caption = lbl31.Caption Or lbl50.Caption = lbl32.Caption Or lbl50.Caption = lbl33.Caption Or lbl50.Caption = lbl40.Caption Or lbl50.Caption = lbl41.Caption Or lbl50.Caption = lbl42.Caption Then
    lbl50.ForeColor = &HFF&
Else
    lbl50.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl51_Click()
lbl51.Caption = cmbNumeros.Text
If lbl51.Caption = lbl21.Caption Or lbl51.Caption = lbl24.Caption Or lbl51.Caption = lbl27.Caption Or lbl51.Caption = lbl75.Caption Or lbl51.Caption = lbl78.Caption Or lbl51.Caption = lbl81.Caption Or lbl51.Caption = lbl46.Caption Or lbl51.Caption = lbl47.Caption Or lbl51.Caption = lbl48.Caption Or lbl51.Caption = lbl49.Caption Or lbl51.Caption = lbl50.Caption Or lbl51.Caption = lbl52.Caption Or lbl51.Caption = lbl53.Caption Or lbl51.Caption = lbl54.Caption Or lbl51.Caption = lbl31.Caption Or lbl51.Caption = lbl32.Caption Or lbl51.Caption = lbl33.Caption Or lbl51.Caption = lbl40.Caption Or lbl51.Caption = lbl41.Caption Or lbl51.Caption = lbl42.Caption Then
    lbl51.ForeColor = &HFF&
Else
    lbl51.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl53_Click()
lbl53.Caption = cmbNumeros.Text
If lbl53.Caption = lbl20.Caption Or lbl53.Caption = lbl23.Caption Or lbl53.Caption = lbl26.Caption Or lbl53.Caption = lbl74.Caption Or lbl53.Caption = lbl77.Caption Or lbl53.Caption = lbl80.Caption Or lbl53.Caption = lbl46.Caption Or lbl53.Caption = lbl47.Caption Or lbl53.Caption = lbl48.Caption Or lbl53.Caption = lbl49.Caption Or lbl53.Caption = lbl50.Caption Or lbl53.Caption = lbl51.Caption Or lbl53.Caption = lbl52.Caption Or lbl53.Caption = lbl54.Caption Or lbl53.Caption = lbl34.Caption Or lbl53.Caption = lbl35.Caption Or lbl53.Caption = lbl36.Caption Or lbl53.Caption = lbl43.Caption Or lbl53.Caption = lbl44.Caption Or lbl53.Caption = lbl45.Caption Then
    lbl53.ForeColor = &HFF&
Else
    lbl53.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl54_Click()
lbl54.Caption = cmbNumeros.Text
If lbl54.Caption = lbl21.Caption Or lbl54.Caption = lbl24.Caption Or lbl54.Caption = lbl27.Caption Or lbl54.Caption = lbl75.Caption Or lbl54.Caption = lbl78.Caption Or lbl54.Caption = lbl81.Caption Or lbl54.Caption = lbl46.Caption Or lbl54.Caption = lbl47.Caption Or lbl54.Caption = lbl48.Caption Or lbl54.Caption = lbl49.Caption Or lbl54.Caption = lbl50.Caption Or lbl54.Caption = lbl51.Caption Or lbl54.Caption = lbl52.Caption Or lbl54.Caption = lbl53.Caption Or lbl54.Caption = lbl34.Caption Or lbl54.Caption = lbl35.Caption Or lbl54.Caption = lbl36.Caption Or lbl54.Caption = lbl43.Caption Or lbl54.Caption = lbl44.Caption Or lbl54.Caption = lbl45.Caption Then
    lbl54.ForeColor = &HFF&
Else
    lbl54.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl55_Click()
lbl55.Caption = cmbNumeros.Text
If lbl55.Caption = lbl1.Caption Or lbl55.Caption = lbl4.Caption Or lbl55.Caption = lbl7.Caption Or lbl55.Caption = lbl28.Caption Or lbl55.Caption = lbl31.Caption Or lbl55.Caption = lbl34.Caption Or lbl55.Caption = lbl56.Caption Or lbl55.Caption = lbl57.Caption Or lbl55.Caption = lbl58.Caption Or lbl55.Caption = lbl59.Caption Or lbl55.Caption = lbl60.Caption Or lbl55.Caption = lbl61.Caption Or lbl55.Caption = lbl62.Caption Or lbl55.Caption = lbl63.Caption Or lbl55.Caption = lbl64.Caption Or lbl55.Caption = lbl65.Caption Or lbl55.Caption = lbl66.Caption Or lbl55.Caption = lbl73.Caption Or lbl55.Caption = lbl74.Caption Or lbl55.Caption = lbl75.Caption Then
    lbl55.ForeColor = &HFF&
Else
    lbl55.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl56_Click()
lbl56.Caption = cmbNumeros.Text
If lbl56.Caption = lbl2.Caption Or lbl56.Caption = lbl5.Caption Or lbl56.Caption = lbl8.Caption Or lbl56.Caption = lbl29.Caption Or lbl56.Caption = lbl32.Caption Or lbl56.Caption = lbl35.Caption Or lbl56.Caption = lbl55.Caption Or lbl56.Caption = lbl57.Caption Or lbl56.Caption = lbl58.Caption Or lbl56.Caption = lbl59.Caption Or lbl56.Caption = lbl60.Caption Or lbl56.Caption = lbl61.Caption Or lbl56.Caption = lbl62.Caption Or lbl56.Caption = lbl63.Caption Or lbl56.Caption = lbl64.Caption Or lbl56.Caption = lbl65.Caption Or lbl56.Caption = lbl66.Caption Or lbl56.Caption = lbl73.Caption Or lbl56.Caption = lbl74.Caption Or lbl56.Caption = lbl75.Caption Then
    lbl56.ForeColor = &HFF&
Else
    lbl56.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl57_Click()
lbl57.Caption = cmbNumeros.Text
If lbl57.Caption = lbl3.Caption Or lbl57.Caption = lbl6.Caption Or lbl57.Caption = lbl9.Caption Or lbl57.Caption = lbl30.Caption Or lbl57.Caption = lbl33.Caption Or lbl57.Caption = lbl36.Caption Or lbl57.Caption = lbl55.Caption Or lbl57.Caption = lbl56.Caption Or lbl57.Caption = lbl58.Caption Or lbl57.Caption = lbl59.Caption Or lbl57.Caption = lbl60.Caption Or lbl57.Caption = lbl61.Caption Or lbl57.Caption = lbl62.Caption Or lbl57.Caption = lbl63.Caption Or lbl57.Caption = lbl64.Caption Or lbl57.Caption = lbl65.Caption Or lbl57.Caption = lbl66.Caption Or lbl57.Caption = lbl73.Caption Or lbl57.Caption = lbl74.Caption Or lbl57.Caption = lbl75.Caption Then
    lbl57.ForeColor = &HFF&
Else
    lbl57.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl58_Click()
lbl58.Caption = cmbNumeros.Text
If lbl58.Caption = lbl1.Caption Or lbl58.Caption = lbl4.Caption Or lbl58.Caption = lbl7.Caption Or lbl58.Caption = lbl28.Caption Or lbl58.Caption = lbl31.Caption Or lbl58.Caption = lbl34.Caption Or lbl58.Caption = lbl55.Caption Or lbl58.Caption = lbl56.Caption Or lbl58.Caption = lbl57.Caption Or lbl58.Caption = lbl59.Caption Or lbl58.Caption = lbl60.Caption Or lbl58.Caption = lbl61.Caption Or lbl58.Caption = lbl62.Caption Or lbl58.Caption = lbl63.Caption Or lbl58.Caption = lbl67.Caption Or lbl58.Caption = lbl68.Caption Or lbl58.Caption = lbl69.Caption Or lbl58.Caption = lbl76.Caption Or lbl58.Caption = lbl77.Caption Or lbl58.Caption = lbl78.Caption Then
    lbl58.ForeColor = &HFF&
Else
    lbl58.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl59_Click()
lbl59.Caption = cmbNumeros.Text
If lbl59.Caption = lbl2.Caption Or lbl59.Caption = lbl5.Caption Or lbl59.Caption = lbl8.Caption Or lbl59.Caption = lbl29.Caption Or lbl59.Caption = lbl32.Caption Or lbl59.Caption = lbl35.Caption Or lbl59.Caption = lbl55.Caption Or lbl59.Caption = lbl56.Caption Or lbl59.Caption = lbl57.Caption Or lbl59.Caption = lbl58.Caption Or lbl59.Caption = lbl60.Caption Or lbl59.Caption = lbl61.Caption Or lbl59.Caption = lbl62.Caption Or lbl59.Caption = lbl63.Caption Or lbl59.Caption = lbl67.Caption Or lbl59.Caption = lbl68.Caption Or lbl59.Caption = lbl69.Caption Or lbl59.Caption = lbl76.Caption Or lbl59.Caption = lbl77.Caption Or lbl59.Caption = lbl78.Caption Then
    lbl59.ForeColor = &HFF&
Else
    lbl59.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl6_Click()
lbl6.Caption = cmbNumeros.Text
If lbl6.Caption = lbl30.Caption Or lbl6.Caption = lbl33.Caption Or lbl6.Caption = lbl36.Caption Or lbl6.Caption = lbl57.Caption Or lbl6.Caption = lbl60.Caption Or lbl6.Caption = lbl63.Caption Or lbl6.Caption = lbl1.Caption Or lbl6.Caption = lbl2.Caption Or lbl6.Caption = lbl3.Caption Or lbl6.Caption = lbl4.Caption Or lbl6.Caption = lbl5.Caption Or lbl6.Caption = lbl7.Caption Or lbl6.Caption = lbl8.Caption Or lbl6.Caption = lbl9.Caption Or lbl6.Caption = lbl13.Caption Or lbl6.Caption = lbl14.Caption Or lbl6.Caption = lbl15.Caption Or lbl6.Caption = lbl22.Caption Or lbl6.Caption = lbl23.Caption Or lbl6.Caption = lbl24.Caption Then
    lbl6.ForeColor = &HFF&
Else
    lbl6.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl60_Click()
lbl60.Caption = cmbNumeros.Text
If lbl60.Caption = lbl3.Caption Or lbl60.Caption = lbl6.Caption Or lbl60.Caption = lbl9.Caption Or lbl60.Caption = lbl30.Caption Or lbl60.Caption = lbl33.Caption Or lbl60.Caption = lbl36.Caption Or lbl60.Caption = lbl55.Caption Or lbl60.Caption = lbl56.Caption Or lbl60.Caption = lbl57.Caption Or lbl60.Caption = lbl58.Caption Or lbl60.Caption = lbl59.Caption Or lbl60.Caption = lbl61.Caption Or lbl60.Caption = lbl62.Caption Or lbl60.Caption = lbl63.Caption Or lbl60.Caption = lbl67.Caption Or lbl60.Caption = lbl68.Caption Or lbl60.Caption = lbl69.Caption Or lbl60.Caption = lbl76.Caption Or lbl60.Caption = lbl77.Caption Or lbl60.Caption = lbl78.Caption Then
    lbl60.ForeColor = &HFF&
Else
    lbl60.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl62_Click()
lbl62.Caption = cmbNumeros.Text
If lbl62.Caption = lbl2.Caption Or lbl62.Caption = lbl5.Caption Or lbl62.Caption = lbl8.Caption Or lbl62.Caption = lbl29.Caption Or lbl62.Caption = lbl32.Caption Or lbl62.Caption = lbl35.Caption Or lbl62.Caption = lbl55.Caption Or lbl62.Caption = lbl56.Caption Or lbl62.Caption = lbl57.Caption Or lbl62.Caption = lbl58.Caption Or lbl62.Caption = lbl59.Caption Or lbl62.Caption = lbl60.Caption Or lbl62.Caption = lbl61.Caption Or lbl62.Caption = lbl63.Caption Or lbl62.Caption = lbl70.Caption Or lbl62.Caption = lbl71.Caption Or lbl62.Caption = lbl72.Caption Or lbl62.Caption = lbl79.Caption Or lbl62.Caption = lbl80.Caption Or lbl62.Caption = lbl81.Caption Then
    lbl62.ForeColor = &HFF&
Else
    lbl62.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl63_Click()
lbl63.Caption = cmbNumeros.Text
If lbl63.Caption = lbl3.Caption Or lbl63.Caption = lbl6.Caption Or lbl63.Caption = lbl9.Caption Or lbl63.Caption = lbl30.Caption Or lbl63.Caption = lbl33.Caption Or lbl63.Caption = lbl36.Caption Or lbl63.Caption = lbl55.Caption Or lbl63.Caption = lbl56.Caption Or lbl63.Caption = lbl57.Caption Or lbl63.Caption = lbl58.Caption Or lbl63.Caption = lbl59.Caption Or lbl63.Caption = lbl60.Caption Or lbl63.Caption = lbl61.Caption Or lbl63.Caption = lbl62.Caption Or lbl63.Caption = lbl70.Caption Or lbl63.Caption = lbl71.Caption Or lbl63.Caption = lbl72.Caption Or lbl63.Caption = lbl79.Caption Or lbl63.Caption = lbl80.Caption Or lbl63.Caption = lbl81.Caption Then
    lbl63.ForeColor = &HFF&
Else
    lbl63.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl64_Click()
lbl64.Caption = cmbNumeros.Text
If lbl64.Caption = lbl10.Caption Or lbl64.Caption = lbl13.Caption Or lbl64.Caption = lbl16.Caption Or lbl64.Caption = lbl37.Caption Or lbl64.Caption = lbl40.Caption Or lbl64.Caption = lbl43.Caption Or lbl64.Caption = lbl65.Caption Or lbl64.Caption = lbl66.Caption Or lbl64.Caption = lbl67.Caption Or lbl64.Caption = lbl68.Caption Or lbl64.Caption = lbl69.Caption Or lbl64.Caption = lbl70.Caption Or lbl64.Caption = lbl71.Caption Or lbl64.Caption = lbl72.Caption Or lbl64.Caption = lbl55.Caption Or lbl64.Caption = lbl56.Caption Or lbl64.Caption = lbl57.Caption Or lbl64.Caption = lbl73.Caption Or lbl64.Caption = lbl74.Caption Or lbl64.Caption = lbl75.Caption Then
    lbl64.ForeColor = &HFF&
Else
    lbl64.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl65_Click()
lbl65.Caption = cmbNumeros.Text
If lbl65.Caption = lbl11.Caption Or lbl65.Caption = lbl14.Caption Or lbl65.Caption = lbl17.Caption Or lbl65.Caption = lbl38.Caption Or lbl65.Caption = lbl41.Caption Or lbl65.Caption = lbl44.Caption Or lbl65.Caption = lbl64.Caption Or lbl65.Caption = lbl66.Caption Or lbl65.Caption = lbl67.Caption Or lbl65.Caption = lbl68.Caption Or lbl65.Caption = lbl69.Caption Or lbl65.Caption = lbl70.Caption Or lbl65.Caption = lbl71.Caption Or lbl65.Caption = lbl72.Caption Or lbl65.Caption = lbl55.Caption Or lbl65.Caption = lbl56.Caption Or lbl65.Caption = lbl57.Caption Or lbl65.Caption = lbl73.Caption Or lbl65.Caption = lbl74.Caption Or lbl65.Caption = lbl75.Caption Then
    lbl65.ForeColor = &HFF&
Else
    lbl65.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl66_Click()
lbl66.Caption = cmbNumeros.Text
If lbl66.Caption = lbl12.Caption Or lbl66.Caption = lbl15.Caption Or lbl66.Caption = lbl18.Caption Or lbl66.Caption = lbl39.Caption Or lbl66.Caption = lbl42.Caption Or lbl66.Caption = lbl45.Caption Or lbl66.Caption = lbl64.Caption Or lbl66.Caption = lbl65.Caption Or lbl66.Caption = lbl67.Caption Or lbl66.Caption = lbl68.Caption Or lbl66.Caption = lbl69.Caption Or lbl66.Caption = lbl70.Caption Or lbl66.Caption = lbl71.Caption Or lbl66.Caption = lbl72.Caption Or lbl66.Caption = lbl55.Caption Or lbl66.Caption = lbl56.Caption Or lbl66.Caption = lbl57.Caption Or lbl66.Caption = lbl73.Caption Or lbl66.Caption = lbl74.Caption Or lbl66.Caption = lbl75.Caption Then
    lbl66.ForeColor = &HFF&
Else
    lbl66.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl67_Click()
lbl67.Caption = cmbNumeros.Text
If lbl67.Caption = lbl10.Caption Or lbl67.Caption = lbl13.Caption Or lbl67.Caption = lbl16.Caption Or lbl67.Caption = lbl37.Caption Or lbl67.Caption = lbl40.Caption Or lbl67.Caption = lbl43.Caption Or lbl67.Caption = lbl64.Caption Or lbl67.Caption = lbl65.Caption Or lbl67.Caption = lbl66.Caption Or lbl67.Caption = lbl68.Caption Or lbl67.Caption = lbl69.Caption Or lbl67.Caption = lbl70.Caption Or lbl67.Caption = lbl71.Caption Or lbl67.Caption = lbl72.Caption Or lbl67.Caption = lbl58.Caption Or lbl67.Caption = lbl59.Caption Or lbl67.Caption = lbl60.Caption Or lbl67.Caption = lbl76.Caption Or lbl67.Caption = lbl77.Caption Or lbl67.Caption = lbl78.Caption Then
    lbl67.ForeColor = &HFF&
Else
    lbl67.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl69_Click()
lbl69.Caption = cmbNumeros.Text
If lbl69.Caption = lbl12.Caption Or lbl69.Caption = lbl15.Caption Or lbl69.Caption = lbl18.Caption Or lbl69.Caption = lbl39.Caption Or lbl69.Caption = lbl42.Caption Or lbl69.Caption = lbl45.Caption Or lbl69.Caption = lbl64.Caption Or lbl69.Caption = lbl65.Caption Or lbl69.Caption = lbl66.Caption Or lbl69.Caption = lbl67.Caption Or lbl69.Caption = lbl68.Caption Or lbl69.Caption = lbl70.Caption Or lbl69.Caption = lbl71.Caption Or lbl69.Caption = lbl72.Caption Or lbl69.Caption = lbl58.Caption Or lbl69.Caption = lbl59.Caption Or lbl69.Caption = lbl60.Caption Or lbl69.Caption = lbl76.Caption Or lbl69.Caption = lbl77.Caption Or lbl69.Caption = lbl78.Caption Then
    lbl69.ForeColor = &HFF&
Else
    lbl69.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl7_Click()
lbl7.Caption = cmbNumeros.Text
If lbl7.Caption = lbl28.Caption Or lbl7.Caption = lbl31.Caption Or lbl7.Caption = lbl34.Caption Or lbl7.Caption = lbl55.Caption Or lbl7.Caption = lbl58.Caption Or lbl7.Caption = lbl61.Caption Or lbl7.Caption = lbl1.Caption Or lbl7.Caption = lbl2.Caption Or lbl7.Caption = lbl3.Caption Or lbl7.Caption = lbl4.Caption Or lbl7.Caption = lbl5.Caption Or lbl7.Caption = lbl6.Caption Or lbl7.Caption = lbl8.Caption Or lbl7.Caption = lbl9.Caption Or lbl7.Caption = lbl16.Caption Or lbl7.Caption = lbl17.Caption Or lbl7.Caption = lbl18.Caption Or lbl7.Caption = lbl25.Caption Or lbl7.Caption = lbl26.Caption Or lbl7.Caption = lbl27.Caption Then
    lbl7.ForeColor = &HFF&
Else
    lbl7.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl70_Click()
lbl70.Caption = cmbNumeros.Text
If lbl70.Caption = lbl10.Caption Or lbl70.Caption = lbl13.Caption Or lbl70.Caption = lbl16.Caption Or lbl70.Caption = lbl37.Caption Or lbl70.Caption = lbl40.Caption Or lbl70.Caption = lbl43.Caption Or lbl70.Caption = lbl64.Caption Or lbl70.Caption = lbl65.Caption Or lbl70.Caption = lbl66.Caption Or lbl70.Caption = lbl67.Caption Or lbl70.Caption = lbl68.Caption Or lbl70.Caption = lbl69.Caption Or lbl70.Caption = lbl71.Caption Or lbl70.Caption = lbl72.Caption Or lbl70.Caption = lbl61.Caption Or lbl70.Caption = lbl62.Caption Or lbl70.Caption = lbl63.Caption Or lbl70.Caption = lbl79.Caption Or lbl70.Caption = lbl80.Caption Or lbl70.Caption = lbl81.Caption Then
    lbl70.ForeColor = &HFF&
Else
    lbl70.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl71_Click()
lbl71.Caption = cmbNumeros.Text
If lbl71.Caption = lbl11.Caption Or lbl71.Caption = lbl14.Caption Or lbl71.Caption = lbl17.Caption Or lbl71.Caption = lbl38.Caption Or lbl71.Caption = lbl41.Caption Or lbl71.Caption = lbl44.Caption Or lbl71.Caption = lbl64.Caption Or lbl71.Caption = lbl65.Caption Or lbl71.Caption = lbl66.Caption Or lbl71.Caption = lbl67.Caption Or lbl71.Caption = lbl68.Caption Or lbl71.Caption = lbl69.Caption Or lbl71.Caption = lbl70.Caption Or lbl71.Caption = lbl72.Caption Or lbl71.Caption = lbl61.Caption Or lbl71.Caption = lbl62.Caption Or lbl71.Caption = lbl63.Caption Or lbl71.Caption = lbl79.Caption Or lbl71.Caption = lbl80.Caption Or lbl71.Caption = lbl81.Caption Then
    lbl71.ForeColor = &HFF&
Else
    lbl71.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl72_Click()
lbl72.Caption = cmbNumeros.Text
If lbl72.Caption = lbl12.Caption Or lbl72.Caption = lbl15.Caption Or lbl72.Caption = lbl18.Caption Or lbl72.Caption = lbl39.Caption Or lbl72.Caption = lbl42.Caption Or lbl72.Caption = lbl45.Caption Or lbl72.Caption = lbl64.Caption Or lbl72.Caption = lbl65.Caption Or lbl72.Caption = lbl66.Caption Or lbl72.Caption = lbl67.Caption Or lbl72.Caption = lbl68.Caption Or lbl72.Caption = lbl69.Caption Or lbl72.Caption = lbl70.Caption Or lbl72.Caption = lbl71.Caption Or lbl72.Caption = lbl61.Caption Or lbl72.Caption = lbl62.Caption Or lbl72.Caption = lbl63.Caption Or lbl72.Caption = lbl79.Caption Or lbl72.Caption = lbl80.Caption Or lbl72.Caption = lbl81.Caption Then
    lbl72.ForeColor = &HFF&
Else
    lbl72.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl73_Click()
lbl73.Caption = cmbNumeros.Text
If lbl73.Caption = lbl19.Caption Or lbl73.Caption = lbl22.Caption Or lbl73.Caption = lbl25.Caption Or lbl73.Caption = lbl46.Caption Or lbl73.Caption = lbl49.Caption Or lbl73.Caption = lbl52.Caption Or lbl73.Caption = lbl74.Caption Or lbl73.Caption = lbl75.Caption Or lbl73.Caption = lbl76.Caption Or lbl73.Caption = lbl77.Caption Or lbl73.Caption = lbl78.Caption Or lbl73.Caption = lbl79.Caption Or lbl73.Caption = lbl80.Caption Or lbl73.Caption = lbl81.Caption Or lbl73.Caption = lbl55.Caption Or lbl73.Caption = lbl56.Caption Or lbl73.Caption = lbl57.Caption Or lbl73.Caption = lbl64.Caption Or lbl73.Caption = lbl65.Caption Or lbl73.Caption = lbl66.Caption Then
    lbl73.ForeColor = &HFF&
Else
    lbl73.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl74_Click()
lbl74.Caption = cmbNumeros.Text
If lbl74.Caption = lbl20.Caption Or lbl74.Caption = lbl23.Caption Or lbl74.Caption = lbl26.Caption Or lbl74.Caption = lbl47.Caption Or lbl74.Caption = lbl50.Caption Or lbl74.Caption = lbl53.Caption Or lbl74.Caption = lbl73.Caption Or lbl74.Caption = lbl75.Caption Or lbl74.Caption = lbl76.Caption Or lbl74.Caption = lbl77.Caption Or lbl74.Caption = lbl78.Caption Or lbl74.Caption = lbl79.Caption Or lbl74.Caption = lbl80.Caption Or lbl74.Caption = lbl81.Caption Or lbl74.Caption = lbl55.Caption Or lbl74.Caption = lbl56.Caption Or lbl74.Caption = lbl57.Caption Or lbl74.Caption = lbl64.Caption Or lbl74.Caption = lbl65.Caption Or lbl74.Caption = lbl66.Caption Then
    lbl74.ForeColor = &HFF&
Else
    lbl74.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl76_Click()
lbl76.Caption = cmbNumeros.Text
If lbl76.Caption = lbl19.Caption Or lbl76.Caption = lbl22.Caption Or lbl76.Caption = lbl25.Caption Or lbl76.Caption = lbl46.Caption Or lbl76.Caption = lbl49.Caption Or lbl76.Caption = lbl52.Caption Or lbl76.Caption = lbl73.Caption Or lbl76.Caption = lbl74.Caption Or lbl76.Caption = lbl75.Caption Or lbl76.Caption = lbl77.Caption Or lbl76.Caption = lbl78.Caption Or lbl76.Caption = lbl79.Caption Or lbl76.Caption = lbl80.Caption Or lbl76.Caption = lbl81.Caption Or lbl76.Caption = lbl58.Caption Or lbl76.Caption = lbl59.Caption Or lbl76.Caption = lbl60.Caption Or lbl76.Caption = lbl67.Caption Or lbl76.Caption = lbl68.Caption Or lbl76.Caption = lbl69.Caption Then
    lbl76.ForeColor = &HFF&
Else
    lbl76.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl77_Click()
lbl77.Caption = cmbNumeros.Text
If lbl77.Caption = lbl20.Caption Or lbl77.Caption = lbl23.Caption Or lbl77.Caption = lbl26.Caption Or lbl77.Caption = lbl47.Caption Or lbl77.Caption = lbl50.Caption Or lbl77.Caption = lbl53.Caption Or lbl77.Caption = lbl73.Caption Or lbl77.Caption = lbl74.Caption Or lbl77.Caption = lbl75.Caption Or lbl77.Caption = lbl76.Caption Or lbl77.Caption = lbl78.Caption Or lbl77.Caption = lbl79.Caption Or lbl77.Caption = lbl80.Caption Or lbl77.Caption = lbl81.Caption Or lbl77.Caption = lbl58.Caption Or lbl77.Caption = lbl59.Caption Or lbl77.Caption = lbl60.Caption Or lbl77.Caption = lbl67.Caption Or lbl77.Caption = lbl68.Caption Or lbl77.Caption = lbl69.Caption Then
    lbl77.ForeColor = &HFF&
Else
    lbl77.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl78_Click()
lbl78.Caption = cmbNumeros.Text
If lbl78.Caption = lbl21.Caption Or lbl78.Caption = lbl24.Caption Or lbl78.Caption = lbl27.Caption Or lbl78.Caption = lbl48.Caption Or lbl78.Caption = lbl51.Caption Or lbl78.Caption = lbl54.Caption Or lbl78.Caption = lbl73.Caption Or lbl78.Caption = lbl74.Caption Or lbl78.Caption = lbl75.Caption Or lbl78.Caption = lbl76.Caption Or lbl78.Caption = lbl77.Caption Or lbl78.Caption = lbl79.Caption Or lbl78.Caption = lbl80.Caption Or lbl78.Caption = lbl81.Caption Or lbl78.Caption = lbl58.Caption Or lbl78.Caption = lbl59.Caption Or lbl78.Caption = lbl60.Caption Or lbl78.Caption = lbl67.Caption Or lbl78.Caption = lbl68.Caption Or lbl78.Caption = lbl69.Caption Then
    lbl78.ForeColor = &HFF&
Else
    lbl78.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl79_Click()
lbl79.Caption = cmbNumeros.Text
If lbl79.Caption = lbl19.Caption Or lbl79.Caption = lbl22.Caption Or lbl79.Caption = lbl25.Caption Or lbl79.Caption = lbl46.Caption Or lbl79.Caption = lbl49.Caption Or lbl79.Caption = lbl52.Caption Or lbl79.Caption = lbl73.Caption Or lbl79.Caption = lbl74.Caption Or lbl79.Caption = lbl75.Caption Or lbl79.Caption = lbl76.Caption Or lbl79.Caption = lbl77.Caption Or lbl79.Caption = lbl78.Caption Or lbl79.Caption = lbl80.Caption Or lbl79.Caption = lbl81.Caption Or lbl79.Caption = lbl61.Caption Or lbl79.Caption = lbl62.Caption Or lbl79.Caption = lbl63.Caption Or lbl79.Caption = lbl70.Caption Or lbl79.Caption = lbl71.Caption Or lbl79.Caption = lbl72.Caption Then
    lbl79.ForeColor = &HFF&
Else
    lbl79.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl8_Click()
lbl8.Caption = cmbNumeros.Text
If lbl8.Caption = lbl29.Caption Or lbl8.Caption = lbl32.Caption Or lbl8.Caption = lbl35.Caption Or lbl8.Caption = lbl56.Caption Or lbl8.Caption = lbl59.Caption Or lbl8.Caption = lbl62.Caption Or lbl8.Caption = lbl1.Caption Or lbl8.Caption = lbl2.Caption Or lbl8.Caption = lbl3.Caption Or lbl8.Caption = lbl4.Caption Or lbl8.Caption = lbl5.Caption Or lbl8.Caption = lbl6.Caption Or lbl8.Caption = lbl7.Caption Or lbl8.Caption = lbl9.Caption Or lbl8.Caption = lbl16.Caption Or lbl8.Caption = lbl17.Caption Or lbl8.Caption = lbl18.Caption Or lbl8.Caption = lbl25.Caption Or lbl8.Caption = lbl26.Caption Or lbl8.Caption = lbl27.Caption Then
    lbl8.ForeColor = &HFF&
Else
    lbl8.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl80_Click()
lbl80.Caption = cmbNumeros.Text
If lbl80.Caption = lbl20.Caption Or lbl80.Caption = lbl23.Caption Or lbl80.Caption = lbl26.Caption Or lbl80.Caption = lbl47.Caption Or lbl80.Caption = lbl50.Caption Or lbl80.Caption = lbl53.Caption Or lbl80.Caption = lbl73.Caption Or lbl80.Caption = lbl74.Caption Or lbl80.Caption = lbl75.Caption Or lbl80.Caption = lbl76.Caption Or lbl80.Caption = lbl77.Caption Or lbl80.Caption = lbl78.Caption Or lbl80.Caption = lbl79.Caption Or lbl80.Caption = lbl81.Caption Or lbl80.Caption = lbl61.Caption Or lbl80.Caption = lbl62.Caption Or lbl80.Caption = lbl63.Caption Or lbl80.Caption = lbl70.Caption Or lbl80.Caption = lbl71.Caption Or lbl80.Caption = lbl72.Caption Then
    lbl80.ForeColor = &HFF&
Else
    lbl80.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl81_Click()
lbl81.Caption = cmbNumeros.Text
If lbl81.Caption = lbl21.Caption Or lbl81.Caption = lbl24.Caption Or lbl81.Caption = lbl27.Caption Or lbl81.Caption = lbl48.Caption Or lbl81.Caption = lbl51.Caption Or lbl81.Caption = lbl54.Caption Or lbl81.Caption = lbl73.Caption Or lbl81.Caption = lbl74.Caption Or lbl81.Caption = lbl75.Caption Or lbl81.Caption = lbl76.Caption Or lbl81.Caption = lbl77.Caption Or lbl81.Caption = lbl78.Caption Or lbl81.Caption = lbl79.Caption Or lbl81.Caption = lbl80.Caption Or lbl81.Caption = lbl61.Caption Or lbl81.Caption = lbl62.Caption Or lbl81.Caption = lbl63.Caption Or lbl81.Caption = lbl70.Caption Or lbl81.Caption = lbl71.Caption Or lbl81.Caption = lbl72.Caption Then
    lbl81.ForeColor = &HFF&
Else
    lbl81.ForeColor = &H0&
End If
    Call Ganhou
End Sub
Private Sub lbl9_Click()
lbl9.Caption = cmbNumeros.Text
If lbl9.Caption = lbl30.Caption Or lbl9.Caption = lbl33.Caption Or lbl9.Caption = lbl36.Caption Or lbl9.Caption = lbl57.Caption Or lbl9.Caption = lbl60.Caption Or lbl9.Caption = lbl63.Caption Or lbl9.Caption = lbl1.Caption Or lbl9.Caption = lbl2.Caption Or lbl9.Caption = lbl3.Caption Or lbl9.Caption = lbl4.Caption Or lbl9.Caption = lbl5.Caption Or lbl9.Caption = lbl6.Caption Or lbl9.Caption = lbl7.Caption Or lbl9.Caption = lbl8.Caption Or lbl9.Caption = lbl16.Caption Or lbl9.Caption = lbl17.Caption Or lbl9.Caption = lbl18.Caption Or lbl9.Caption = lbl25.Caption Or lbl9.Caption = lbl26.Caption Or lbl9.Caption = lbl27.Caption Then
    lbl9.ForeColor = &HFF&
Else
    lbl9.ForeColor = &H0&
End If
    Call Ganhou
End Sub
