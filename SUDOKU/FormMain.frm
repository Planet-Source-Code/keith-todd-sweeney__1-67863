VERSION 5.00
Begin VB.Form FormMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "                                               SUDOKU  PUZZLE"
   ClientHeight    =   9885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12090
   BeginProperty Font 
      Name            =   "RockFont"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormMain.frx":0ECA
   ScaleHeight     =   9885
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1680
      Top             =   9360
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1080
      Top             =   9360
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Color 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      Picture         =   "FormMain.frx":32E14C
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Color 1"
      DisabledPicture =   "FormMain.frx":3328B2
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Picture         =   "FormMain.frx":336F02
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   990
      Left            =   600
      Top             =   9360
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   2640
      Picture         =   "FormMain.frx":342E9B
      ScaleHeight     =   1455
      ScaleWidth      =   8775
      TabIndex        =   82
      Top             =   8040
      Width           =   8775
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   80
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   79
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   78
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   77
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   76
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   75
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   74
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   73
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   72
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   71
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   70
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   69
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   68
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   67
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   66
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   65
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   64
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   63
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   62
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   61
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   60
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   59
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   58
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   57
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   56
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   55
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   54
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   53
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   52
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   51
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   50
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   49
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   48
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   47
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   46
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   45
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   44
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   43
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   42
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   41
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   40
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   39
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   38
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   37
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   36
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   35
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   34
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   33
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   32
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   31
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   30
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   29
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   28
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   27
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   26
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   25
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   24
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   23
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   22
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   21
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   20
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   19
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   18
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   17
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "EnglishTowne-Normal"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   9120
      TabIndex        =   87
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "EnglishTowne-Normal"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   9120
      TabIndex        =   86
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "EnglishTowne-Normal"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   120
      TabIndex        =   85
      Top             =   240
      Width           =   2655
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   7080
      X2              =   7080
      Y1              =   0
      Y2              =   5880
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5040
      X2              =   5040
      Y1              =   0
      Y2              =   5880
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   81
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   3120
      X2              =   9000
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   3120
      X2              =   9000
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Public CurrentIndex As Integer
Public Currentpic As Integer
Public Picnum As Integer
Public Whatstr As String
Public Nstr As Integer
Public Start As Boolean
Public Finishok As Boolean
Public T As String
Public T1 As String
Public T2 As String
Public picfile As String
Public HourDiff As Integer
Public MinDiff As Integer
Public SecDiff As Integer
Public Q As Integer
Private Sub Command2_Click()
 End
End Sub

Sub Form_Load()
Blank.WindowState = 2
Blank.Show

FormMain.Show
FormMain.WindowState = 0
FormMain.Move (Screen.Width - FormMain.Width) / 2, (Screen.Height - FormMain.Height) / 2
Picture2 = LoadPicture(App.Path & "\Buttons2.bmp")

Label2.Caption = "Left Click Square required Then Left Click Number wanted"
Start = True
Filenum = FreeFile
If OldFile Then
   Open App.Path & "\Puzzles\Saved Files\" & Infile For Input As #Filenum
   Infile = Sname
Else
   Open App.Path & "\Puzzles\" & Diff & "\" & Infile & ".txt" For Input As #Filenum
End If
K = 0
For I = 1 To 81 Step 9
   For J = 1 To 9
     Input #Filenum, Snum
     Command1(J + I - 2).FontBold = True
     SavePuzz(K) = Snum
     K = K + 1
     If Snum = 0 Then
        Command1(J + I - 2).Caption = ""
        Command1(J + I - 2).Picture = Command3.Picture
     Else
        Command1(J + I - 2).Picture = Command2.Picture
        Command1(J + I - 2).Font = "Igloo"
        Command1(J + I - 2).FontSize = 26
        Command1(J + I - 2).Caption = Snum
        Pos(J + I - 2) = True
     End If
   Next J
Next I
Close #Filenum
Cont:
'castle.Visible = True
'castle.Show
End Sub
Public Function IsMouseOver(hWnd As Long) As Boolean
    Dim Mouse As POINTAPI
    GetCursorPos Mouse
    I = Mouse.X
    J = Mouse.Y
    If J < 650 And J > 610 Then ' Bottom row of buttons
       If I < 410 And I > 330 Then 'EXIT
         Whatstr = "EXIT"
       End If
       If I < 507 And I > 447 Then 'Get Solution
         Whatstr = "GETSOLUTION"
       End If
       If I < 650 And I > 580 Then 'Save Puzzle working on
          Whatstr = "SAVEPUZZLE"
       End If
       If I < 810 And I > 750 Then 'CHEAT
         Whatstr = "CHEAT"
       End If
    End If
End Function
Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     PositionBefore = PositionNow
     Call api.GetCursorPos(PositionNow)
     Try = IsMouseOver(hWnd)
     If Button = 1 Then
       Nstr2 = Nstr
     End If
End Sub
Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     PositionBefore = PositionNow
     Call api.GetCursorPos(PositionNow)
     Try = IsMouseOver(hWnd)
     If Button = 1 Then
       If Whatstr = "EXIT" Then
         End
       End If
       If Whatstr = "GETSOLUTION" Then
         Call GetSolution
       End If
       If Whatstr = "SAVEPUZZLE" Then
         Call SaveFile
       End If
       If Whatstr = "CHEAT" Then
         Call Cheat
       End If
    End If
End Sub
Sub GetSolution()
  Timer1.Enabled = False
  Filenum = FreeFile
  Open App.Path & "\Puzzles\Solutions\" & Infile & ".txt" For Input As #Filenum
  For I = 1 To 81 Step 9
     For J = 1 To 9
      Input #Filenum, Snum
      Command1(J + I - 2).Font = "Igloo"
      Command1(J + I - 2).Picture = Command3.Picture
      Command1(J + I - 2).FontSize = 24
      Command1(J + I - 2).FontBold = True
      Command1(J + I - 2).Caption = Snum
    Next J
  Next I
  Close #Filenum
End Sub
Sub SaveFile()
  Timer1.Enabled = False
  Filenum = FreeFile
  Open App.Path & "\Puzzles\Saved Files\" & Infile & "Saved" & ".txt" For Output As #Filenum
  SFile = Infile & "Saved" & ".txt"
  K = 0
  For I = 1 To 81 Step 9
     For J = 1 To 9
      Print #Filenum, SavePuzz(K);
      K = K + 1
    Next J
  Next I
  Close #Filenum
  MsgBox "File Name of saved File is " & SFile & "Please take note of filename "
  End Sub


Sub Cheat()
'Check Numbers placed on puzzle with numbers stored in solution
Filenum = FreeFile
  Open App.Path & "\Puzzles\Solutions\" & Infile & ".txt" For Input As #Filenum
  For I = 1 To 81 Step 9
     For J = 1 To 9
      Input #Filenum, Snum
      If Command1(J + I - 2).Caption = "" Then
         GoTo A:
      End If
      Temp = Val(Command1(J + I - 2).Caption)
      If Snum <> Val(Command1(J + I - 2).Caption) Then
        For K = 1 To 4
         Command1(J + I - 2).Caption = ""
         Call Wait(0.2)
         Beep
         Command1(J + I - 2).Caption = Temp
         Call Wait(0.2)
        Next
      Command1(J + I - 2).Caption = ""
      End If
A:    Next J
  Next I
  Close #Filenum
End Sub
Sub Finished()
Filenum = FreeFile
  Open App.Path & "\Puzzles\Solutions\" & Infile & ".txt" For Input As #Filenum
  Tot = 0
  For I = 1 To 81 Step 9
     For J = 1 To 9
      Input #Filenum, Snum
      If Command1(J + I - 2).Caption = "" Then
         GoTo A:
      End If
      If Snum = Val(Command1(J + I - 2).Caption) Then
        Tot = Tot + 1
      End If
      If Tot = 81 Then
        Finishok = True
      End If
A:    Next J
  Next I
  Close #Filenum
End Sub
Sub Over()
 Call PlaySnd(Done)
 Call Wait(1)
 Label2.Caption = ""
 Timer1.Enabled = False
 Timer2.Enabled = False
 Label2.Caption = "Well done you have completed the puzzle"
End Sub
Private Sub Timer1_Timer()
 'Display elapsed time on form
 Timer1.Enabled = True
 If Start = True Then
    T = Time
    Start = False
 End If
 T1 = Time
'Difference between start time (T)and current time T1
 Diff = TimeValue(T1) - TimeValue(T)
 T2 = Format(Diff, "hh:mm:ss")
 Label4.Caption = T2
End Sub

Private Sub Wait(interval) 'Wait for specified interval(1 = 1 second)
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

'**
'Putting numbers into squares of puzzle
'**
Private Sub Command1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Mouse As POINTAPI
Dim I As Integer

Label1.Caption = ""
If Button = vbLeftButton Then
   CurrentIndex = Index
   If Command1(CurrentIndex).Caption <> "" And Pos(CurrentIndex) = True Then
       Label1.Visible = True
       Label1.Font = "Igloo"
       Label1.FontSize = 20
       Label1.ForeColor = vbGreen
       Label1.Caption = " Square Fixed"
       Beep
       Call Wait(2)
       Beep
       Label1.Caption = ""
       Label1.Visible = False
   Else
       Currentpic = Int(Index / 9) + 1
       If Currentpic >= 1 And Currentpic < 4 Then Picnum = 1
       If Currentpic >= 4 And Currentpic < 7 Then Picnum = 2
       If Currentpic >= 7 And Currentpic <= 9 Then Picnum = 3
       For I = 0 To 8
            NumForm2.Label1(I).ForeColor = vbGreen
       Next
       NumForm2.Picture = LoadPicture(App.Path & "\Psycho" & Picnum & ".bmp")
       NumForm2.Show
       
       With NumForm2
         .Left = FormMain.Command1(Index).Left - 1500
         .Top = Command1(Index).Top - (Command1(Index).Height / 2) + 400
        If .Top < 800 Then
           .Top = 800
         End If
         If .Top > 5000 Then
           .Top = 4400
         End If
         If Not .Visible Then .Visible = True
       End With
       Call Finished
       If Finishok Then Call Over
     End If
 Else 'clear square
     If Button = vbRightButton Then
         If Pos(CurrentIndex) = True Then
          'do nothing
         Else
           Command1(CurrentIndex).Caption = ""
         End If
     End If
     If NumForm2.Visible Then NumForm2.Visible = False
 End If
End Sub

'
