VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin Project1.activebar activebar2 
      Height          =   2655
      Left            =   2880
      TabIndex        =   11
      Top             =   1920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   4683
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   480
         Width           =   1575
      End
   End
   Begin Project1.flatbutton flatbutton1 
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      AlwaysHighlighted=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.flatbutton restore 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.activebar activebar1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   8070
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Caption         =   "Active Container"
         Height          =   2655
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1935
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   2160
            Width           =   1695
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Check2"
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   1440
            Width           =   915
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Option3"
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   1080
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub flatbutton1_Click(ByVal ClickReason As ClickReason)
activebar2.restore
End Sub

Private Sub restore_Click(ByVal ClickReason As ClickReason)
activebar1.restore
End Sub
