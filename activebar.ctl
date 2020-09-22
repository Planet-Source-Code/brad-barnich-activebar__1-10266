VERSION 5.00
Begin VB.UserControl activebar 
   Alignable       =   -1  'True
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2850
   ControlContainer=   -1  'True
   ForwardFocus    =   -1  'True
   ScaleHeight     =   3240
   ScaleWidth      =   2850
   ToolboxBitmap   =   "activebar.ctx":0000
   Begin Project1.flatbutton closebar 
      Height          =   200
      Left            =   2400
      TabIndex        =   0
      Top             =   65
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      AlwaysHighlighted=   -1  'True
      Caption         =   "r"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   6.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line top2 
      BorderColor     =   &H00EFEFEF&
      X1              =   1680
      X2              =   2400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line lside2 
      BorderColor     =   &H00EFEFEF&
      X1              =   600
      X2              =   600
      Y1              =   720
      Y2              =   1320
   End
   Begin VB.Line lside1 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   500
   End
   Begin VB.Line top1 
      BorderColor     =   &H00E0E0E0&
      X1              =   1680
      X2              =   2400
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line rside2 
      BorderColor     =   &H00CFCFCF&
      X1              =   2400
      X2              =   2400
      Y1              =   2160
      Y2              =   2880
   End
   Begin VB.Line bot1 
      BorderColor     =   &H009F9F9F&
      X1              =   120
      X2              =   960
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line rside1 
      BorderColor     =   &H009F9F9F&
      X1              =   2520
      X2              =   2520
      Y1              =   2160
      Y2              =   2880
   End
   Begin VB.Shape border 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   0
      Top             =   50
      Width           =   2655
   End
   Begin VB.Line bot2 
      BorderColor     =   &H00CFCFCF&
      X1              =   960
      X2              =   120
      Y1              =   2640
      Y2              =   2640
   End
End
Attribute VB_Name = "activebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim X
Dim Y




Private Sub closebar_Click(ByVal ClickReason As ClickReason)
Y = UserControl.Height
X = UserControl.Width
UserControl.Size 25, 0

End Sub

Sub restore()
If UserControl.Height <> 0 Then X = UserControl.Width
If UserControl.Height <> 0 Then Y = UserControl.Height
UserControl.Height = Y
UserControl.Width = X
End Sub

Private Sub UserControl_Resize()
border.BackColor = vbActiveTitleBar
closebar.Left = UserControl.Width - 245
border.Width = UserControl.Width - 20
top1.X1 = 0
top1.Y1 = 0
top1.X2 = UserControl.Width
top1.Y2 = 0
top2.X1 = 15
top2.Y1 = 15
top2.X2 = UserControl.Width - 10
top2.Y2 = 15
bot1.X1 = 0
bot1.Y1 = UserControl.Height - 10
bot1.X2 = UserControl.Width
bot1.Y2 = UserControl.Height - 10
bot2.X1 = 15
bot2.Y1 = UserControl.Height - 25
bot2.X2 = UserControl.Height - 25
bot2.Y2 = UserControl.Height - 25
rside1.X1 = UserControl.Width - 10
rside1.Y1 = 0
rside1.X2 = UserControl.Width - 10
rside1.Y2 = UserControl.Height
rside2.X1 = UserControl.Width - 25
rside2.Y1 = 10
rside2.X2 = UserControl.Width - 25
rside2.Y2 = UserControl.Height - 20
lside1.X1 = 0
lside1.Y1 = 0
lside1.X2 = 0
lside1.Y2 = UserControl.Height
lside2.X1 = 15
lside2.Y1 = 15
lside2.X2 = 15
lside2.Y2 = UserControl.Height - 15
End Sub


