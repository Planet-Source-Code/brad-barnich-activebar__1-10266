VERSION 5.00
Begin VB.UserControl flatbutton 
   Alignable       =   -1  'True
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   ControlContainer=   -1  'True
   ScaleHeight     =   2970
   ScaleWidth      =   2595
   ToolboxBitmap   =   "flatbutton.ctx":0000
   Begin VB.Timer tmrhighlight 
      Left            =   720
      Top             =   1440
   End
   Begin VB.Label lblcaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   660
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "flatbutton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Enum htWhatToApply
    apyDrawBorder = 1
    apyBackColor = 2
    apyCaption = 4
    apyEnabled = 8
    apyFont = 16
    apyAll = (apyBackColor Or apyCaption Or apyEnabled Or apyFont)
End Enum
Dim mbHasCapture As Boolean
Dim mpntLabelPos As POINTAPI
Dim mpntOldSize As POINTAPI

Private Type POINTAPI
    X As Long
    Y As Long
    End Type


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type
    Private Const BDR_RAISEDINNER = &H4
    Private Const BDR_RAISEDOUTER = &H1
    Private Const BDR_SUNKENINNER = &H8
    Private Const BDR_SUNKENOUTER = &H2
    Private Const BDR_MOUSEOVER = BDR_RAISEDINNER
    Private Const BDR_MOUSEDOWN = BDR_SUNKENOUTER
    Private Const BF_BOTTOM = &H8
    Private Const BF_FLAT = &H4000
    Private Const BF_LEFT = &H1
    Private Const BF_RIGHT = &H4
    Private Const BF_TOP = &H2
    Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)


Private Declare Function apiDrawEdge Lib "user32" _
    Alias "DrawEdge" _
    (ByVal hdc As Long, _
    ByRef qrc As RECT, _
    ByVal edge As Long, _
    ByVal grfFlags As Long) As Long
    


Private Declare Function apiGetCursorPos Lib "user32" _
    Alias "GetCursorPos" _
    (lpPoint As POINTAPI) As Long
    


Private Declare Function apiWindowFromPoint Lib "user32" _
    Alias "WindowFromPoint" _
    (ByVal xPoint As Long, _
    ByVal yPoint As Long) As Long
    


Private Declare Function apiDrawFocusRect Lib "user32" _
    Alias "DrawFocusRect" _
    (ByVal hdc As Long, _
    lpRect As RECT) As Long
   
    Private mProp_AlwaysHighlighted As Boolean
    Private mProp_BackColor As OLE_COLOR
    Private mProp_Caption As String
    Private mProp_Enabled As Boolean
    Private mProp_FocusRect As Boolean
    Private mProp_Font As StdFont
    Private mProp_HoverColor As OLE_COLOR
    Const mDef_AlwaysHighlighted = False
    Const mDef_BackColor = vbButtonFace
    Const mDef_Caption = "Flatbutton"
    Const mDef_Enabled = True
    Const mDef_FocusRect = False
    Const mDef_Font = Null ' Ambient.Font
    Const mDef_HoverColor = vbButtonText


Public Enum ClickReason
    ReasonMouse
    ReasonAccessKey
    ReasonKeyboard
End Enum

Event Click(ByVal ClickReason As ClickReason)


Private Sub tmrHighlight_Timer()
    Dim pntCursor As POINTAPI
    apiGetCursorPos pntCursor


    If apiWindowFromPoint(pntCursor.X, pntCursor.Y) = hWnd Then


        If Not mbHasCapture Then
            Call ApplyProperties(apyDrawBorder)
            lblcaption.ForeColor = mProp_HoverColor
            mbHasCapture = True
        End If
    Else


        If mbHasCapture Then
            Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), mProp_BackColor, B
            lblcaption.ForeColor = vbButtonText
            mbHasCapture = False
        End If
    End If
End Sub


Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click(ReasonAccessKey)
End Sub


Private Sub UserControl_Click()
    RaiseEvent Click(ReasonMouse)
End Sub


Private Sub UserControl_EnterFocus()
    Dim rctFocus As RECT
    If Not mProp_FocusRect Then Exit Sub
    rctFocus.Left = 3
    rctFocus.Top = 3
    rctFocus.Right = ScaleWidth - 3
    rctFocus.Bottom = ScaleHeight - 3
    apiDrawFocusRect hdc, rctFocus
    Refresh
End Sub


Private Sub UserControl_ExitFocus()
    If mProp_FocusRect Then Line (3, 3)-(ScaleWidth - 4, ScaleHeight - 4), mProp_BackColor, B
End Sub


Private Sub UserControl_Initialize()
    AutoRedraw = True
    ScaleMode = vbPixels
    lblcaption.Alignment = vbCenter
    lblcaption.AutoSize = True
    lblcaption.BackStyle = vbTransparent
    tmrhighlight.Enabled = False
    tmrhighlight.Interval = 1
End Sub


Private Sub UserControl_InitProperties()
    Width = 1215
    Height = 375
    mProp_AlwaysHighlighted = mDef_AlwaysHighlighted
    mProp_BackColor = mDef_BackColor
    mProp_Caption = mDef_Caption
    mProp_Enabled = mDef_Enabled
    mProp_FocusRect = mDef_FocusRect
    Set mProp_Font = Ambient.Font
    mProp_HoverColor = mDef_HoverColor
    Call ApplyProperties(apyAll)
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mProp_AlwaysHighlighted = PropBag.ReadProperty("AlwaysHighlighted", mDef_AlwaysHighlighted)
    mProp_BackColor = PropBag.ReadProperty("BackColor", mDef_BackColor)
    mProp_Caption = PropBag.ReadProperty("Caption", mDef_Caption)
    mProp_Enabled = PropBag.ReadProperty("Enabled", mDef_Enabled)
    mProp_FocusRect = PropBag.ReadProperty("FocusRect", mDef_FocusRect)
    Set mProp_Font = PropBag.ReadProperty("Font", Ambient.Font)
    mProp_HoverColor = PropBag.ReadProperty("HoverColor", mDef_HoverColor)
    
    Call ApplyProperties(apyAll)


    If Ambient.UserMode Then


        If mProp_AlwaysHighlighted Then
            Call ApplyProperties(apyDrawBorder)
        Else
            tmrhighlight = True
        End If
    End If
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)


    With PropBag
        .WriteProperty "AlwaysHighlighted", mProp_AlwaysHighlighted, mDef_AlwaysHighlighted
        .WriteProperty "BackColor", mProp_BackColor, mDef_BackColor
        .WriteProperty "Caption", mProp_Caption, mDef_Caption
        .WriteProperty "Enabled", mProp_Enabled, mDef_Enabled
        .WriteProperty "FocusRect", mProp_FocusRect, mDef_FocusRect
        .WriteProperty "Font", mProp_Font, Ambient.Font
        .WriteProperty "HoverColor", mProp_HoverColor, mDef_HoverColor
    End With
End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)


    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        UserControl_MouseDown -2, -2, -2, -2
    End If
End Sub


Private Sub UserControl_KeyPress(KeyAscii As Integer)


    If KeyAscii = vbKeySpace Or KeyAscii = vbKeyReturn Then
        RaiseEvent Click(ReasonKeyboard)
    End If
End Sub


Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)


    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        UserControl_MouseUp -2, -2, -2, -2
    End If
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rctBtn As RECT
    Dim dwRetVal As Long
    tmrhighlight.Enabled = False
    lblcaption.Left = mpntLabelPos.X + 1
    lblcaption.Top = mpntLabelPos.Y + 1
    Line (0, 0)-(Width, Height), mProp_BackColor, B
    rctBtn.Left = 0
    rctBtn.Top = 0
    rctBtn.Right = ScaleWidth
    rctBtn.Bottom = ScaleHeight
    dwRetVal = apiDrawEdge(hdc, rctBtn, BDR_MOUSEDOWN, BF_RECT)
End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pntCursor As POINTAPI
    lblcaption.Left = mpntLabelPos.X
    lblcaption.Top = mpntLabelPos.Y
    apiGetCursorPos pntCursor


    If apiWindowFromPoint(pntCursor.X, pntCursor.Y) = hWnd Or mProp_AlwaysHighlighted Then
        Call ApplyProperties(apyDrawBorder)
        mbHasCapture = True
    Else
        Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), mProp_BackColor, B
        mbHasCapture = False
    End If
    If Not mProp_AlwaysHighlighted Then tmrhighlight.Enabled = True
End Sub


Private Sub lblCaption_Click()
    RaiseEvent Click(ReasonMouse)
End Sub


Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, -1, -1
End Sub


Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, -1, -1
End Sub


Private Sub UserControl_Resize()
    Dim rctBtn As RECT
    Dim dwRetVal As Long
    Static sbFirstTime As Boolean


    If Not sbFirstTime Then
        sbFirstTime = True
    Else
        Cls
    End If
    lblcaption.AutoSize = False
    lblcaption.Top = (ScaleHeight / 2) - (lblcaption.Height / 2)
    lblcaption.Left = 1
    lblcaption.Width = ScaleWidth - 2
    


    If Not Ambient.UserMode Or mProp_AlwaysHighlighted Then
        Call ApplyProperties(apyDrawBorder)
    End If
    mpntLabelPos.X = lblcaption.Left
    mpntLabelPos.Y = lblcaption.Top
    mpntOldSize.X = ScaleWidth
    mpntOldSize.Y = ScaleHeight
End Sub
' Private Procedures
' ******************


Private Sub ApplyProperties(ByVal apyWhatToApply As htWhatToApply)
    Dim rctBtn As RECT
    Dim dwRetVal As Long
    Dim n As Long
    If (apyWhatToApply And apyBackColor) Then UserControl.BackColor = mProp_BackColor


    If (apyWhatToApply And apyCaption) Then
        lblcaption.Caption = mProp_Caption
        AccessKeys = ""


        For n = Len(mProp_Caption) To 1 Step -1


            If Mid$(mProp_Caption, n, 1) = "&" Then


                If n = 1 Then
                    AccessKeys = Mid$(mProp_Caption, n + 1, 1)
                ElseIf Not Mid$(mProp_Caption, n - 1, 1) = "&" Then
                    AccessKeys = Mid$(mProp_Caption, n + 1, 1)
                    Exit For
                Else
                    n = n - 1
                End If
            End If
        Next n
    End If


    If (apyWhatToApply And apyFont) Then
        Set UserControl.Font = mProp_Font
        lblcaption.AutoSize = True
        Set lblcaption.Font = mProp_Font
        lblcaption.AutoSize = False
        lblcaption.Top = (ScaleHeight / 2) - (lblcaption.Height / 2)
        lblcaption.Left = 1
        lblcaption.Width = ScaleWidth - 2
    End If
    


    If (apyWhatToApply And apyEnabled) Then


        If Ambient.UserMode Then
            lblcaption.Enabled = mProp_Enabled
            UserControl.Enabled = mProp_Enabled
        End If
    End If
    


    If (apyWhatToApply And apyDrawBorder) Then
        Line (0, 0)-(Width, Height), mProp_BackColor, B
        rctBtn.Left = 0
        rctBtn.Top = 0
        rctBtn.Right = ScaleWidth
        rctBtn.Bottom = ScaleHeight
        
        dwRetVal = apiDrawEdge(hdc, rctBtn, BDR_MOUSEOVER, BF_RECT)
    End If
End Sub
' Properies
' *********


Public Property Get AlwaysHighlighted() As Boolean
    AlwaysHighlighted = mProp_AlwaysHighlighted
End Property


Public Property Let AlwaysHighlighted(ByVal bNewValue As Boolean)


    If Ambient.UserMode Then
        Err.Raise 383
    Else
        mProp_AlwaysHighlighted = bNewValue
        PropertyChanged "AlwaysHighlighted"
    End If
End Property


Public Property Get BackColor() As OLE_COLOR
    BackColor = mProp_BackColor
End Property


Public Property Let BackColor(ByVal oleNewValue As OLE_COLOR)
    mProp_BackColor = oleNewValue
    Call ApplyProperties(apyBackColor Or apyDrawBorder)
    PropertyChanged "BackColor"
End Property


Public Property Get Caption() As String
    Caption = mProp_Caption
End Property


Public Property Let Caption(ByVal sNewValue As String)
    mProp_Caption = sNewValue
    Call ApplyProperties(apyCaption)
    PropertyChanged "Caption"
End Property


Public Property Get FocusRect() As Boolean
    FocusRect = mProp_FocusRect
End Property


Public Property Let FocusRect(ByVal bNewValue As Boolean)


    If Ambient.UserMode Then
        Err.Raise 383
    Else
        mProp_FocusRect = bNewValue
        PropertyChanged "FocusRect"
    End If
End Property


Public Property Get Font() As StdFont
    Set Font = mProp_Font
End Property
Public Property Set Font(ByVal fntNewValue As StdFont)
Set mProp_Font = fntNewValue
Call ApplyProperties(apyFont)
PropertyChanged "Font"
End Property


Public Property Get Enabled() As Boolean
    Enabled = mProp_Enabled
End Property


Public Property Let Enabled(ByVal bNewValue As Boolean)
    mProp_Enabled = bNewValue
    Call ApplyProperties(apyEnabled)
    PropertyChanged "Enabled"
End Property


Public Property Get HoverColor() As OLE_COLOR
    HoverColor = mProp_HoverColor
End Property


Public Property Let HoverColor(ByVal oleNewValue As OLE_COLOR)
    mProp_HoverColor = oleNewValue
    PropertyChanged "HoverColor"
End Property


