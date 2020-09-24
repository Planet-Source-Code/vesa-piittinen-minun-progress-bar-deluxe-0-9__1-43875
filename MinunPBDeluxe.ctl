VERSION 5.00
Begin VB.UserControl MinunPBDeluxe 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   KeyPreview      =   -1  'True
   ScaleHeight     =   129
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   ToolboxBitmap   =   "MinunPBDeluxe.ctx":0000
   Begin VB.PictureBox ShadowMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.PictureBox Shadow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1320
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.PictureBox BarMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.PictureBox Bar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   3375
      Visible         =   0   'False
   End
   Begin VB.PictureBox Corner 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   240
      Visible         =   0   'False
   End
   Begin VB.PictureBox BG 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3375
      Begin VB.PictureBox SG 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   209
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   3135
         Begin VB.PictureBox FG 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   193
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   2895
         End
      End
   End
End
Attribute VB_Name = "MinunPBDeluxe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Minun Progress Bar Deluxe version 0.9
'Please note this is not a finished control, but a work in progress!

'If you make changes, change UserControl name, project name...everything!
'Just to make sure your modified control doesn't mess up with the original :)
'It'll be best for both. If you find a "critical" bug, please mail me.

'The control is freeware, but it'd be nice to know if you spread your own
'version that is based on this :) It'd be also nice to know if you use the
'control in your project(s).

'Last update on 9th of March 2003 - Merri, merry@mbnet.fi
'(original project started in 14th November 2002)
'(deluxe project started in 8th March 2003)
Option Explicit
Public Enum mpbdBarStyle
    mpbdDefault
    mpbdSquares
    mpbdCircles
    mpbdTriangles
    mpbdTrianglesFlipped
End Enum
Public Enum mpbdBorderStyle
    mpbdNormal
    mpbdRoundOne
    mpbdRoundTwo
End Enum
Public Enum mpbdBorderWidth
    mpbdNone
    mpbdOne
    mpbdTwo
    mpbdThree
    mpbdFour
    mpbdFive
    mpbdSix
    mpbdSeven
    mpbdEight
End Enum
Public Enum mpbdColors
    mpbdOneColor
    mpbdTwoColors
    mpbdThreeColors
    mpbdFourColors
End Enum
Public Enum mpbdDirection
    mpbdRight
    mpbdUp
    mpbdLeft
    mpbdDown
End Enum
Public Enum mpbdPercentAlign
    mpbdCenter
    mpbdBarOut
    mpbdBarIn
    mpbdCenterOut
    mpbdCenterIn
    mpbdLeft
    mpbdRight
End Enum
'Default Property Values:
Const m_def_BackAutoColor = 0
Const m_def_BackColor = vbWhite
Const m_def_BackColor2 = &HACD4AC
Const m_def_BackColor3 = &H18000
Const m_def_BackColor4 = &H1FF00
Const m_def_BackColors = 3
Const m_def_BarAutoColor = 0
Const m_def_BarColor = &HD03D30
Const m_def_BarColor2 = &H9C2E24
Const m_def_BarColor3 = &H82261E
Const m_def_BarColor4 = &H681E18
Const m_def_BarColors = 3
Const m_def_BarStyle = 0
Const m_def_BarStyleFullOnly = 0
Const m_def_BorderColor = vbBlack
Const m_def_BorderStyle = 2
Const m_def_BorderWidth = 3
Const m_def_Custom = 0
Const m_def_CustomText = "Done!"
Const m_def_Decimals = 0
Const m_def_Direction = 0
Const m_def_FormColor = &H8000000F
Const m_def_ManualRefresh = False
Const m_def_Max = 10000
Const m_def_Min = 0
Const m_def_NoPercent = False
Const m_def_Percent = 0
Const m_def_PercentAfter = "%"
Const m_def_PercentAlign = 0
Const m_def_PercentBefore = ""
Const m_def_PercentColorBar = &HFFFFFF
Const m_def_PercentColorShadow = &H0
Const m_def_ScaleMode = vbTwips
Const m_def_ShadowColor = &HFCF5F4
Const m_def_ShadowColor2 = &HF9E9E7
Const m_def_ShadowColor3 = &HF3D1CD
Const m_def_ShadowColor4 = &HEDB8B3
'Property Variables:
Dim m_BackAutoColor As Boolean
Dim m_BackColor As Long
Dim m_BackColor2 As Long
Dim m_BackColor3 As Long
Dim m_BackColor4 As Long
Dim m_BackColors As Byte
Dim m_BarAutoColor As Boolean
Dim m_BarColor As Long
Dim m_BarColor2 As Long
Dim m_BarColor3 As Long
Dim m_BarColor4 As Long
Dim m_BarColors As Byte
Dim m_BarStyle As Byte
Dim m_BarStyleFullOnly As Boolean
Dim m_BorderColor As Long
Dim m_BorderStyle As Byte
Dim m_BorderWidth As Byte
Dim m_Custom As Boolean
Dim m_CustomText As String
Dim m_Decimals As Byte
Dim m_Direction As Byte
Dim m_Font As Font
Dim m_FormColor As Long
Dim m_ManualRefresh As Boolean
Dim m_Max As Currency
Dim m_Min As Currency
Dim m_NoPercent As Boolean
Dim m_Percent As Single
Dim m_PercentAfter As String
Dim m_PercentAlign As Byte
Dim m_PercentBefore As String
Dim m_PercentColorBar As Long
Dim m_PercentColorShadow As Long
Dim m_ScaleMode As Integer
Dim m_ShadowColor As Long
Dim m_ShadowColor2 As Long
Dim m_ShadowColor3 As Long
Dim m_ShadowColor4 As Long
Dim m_Value As Currency
'Internal Variables:
Dim m_OldPercent As Single
Dim m_ScaleHeight As Integer
Dim m_ScaleWidth As Integer
Dim m_TextHeight As Integer
Dim m_TextWidth As Integer
Dim Text As String
'API Declarations:
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowParent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Event Declarations:
Event Change()
Attribute Change.VB_Description = "Occurs when the percentage of a control has changed."
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Event Resize()
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Public Property Get BackAutoColor() As Boolean
Attribute BackAutoColor.VB_Description = "Returns/sets the setting used to determine if background colors are automatically calculated."
Attribute BackAutoColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackAutoColor = m_BackAutoColor
End Property
Public Property Let BackAutoColor(ByVal New_BackAutoColor As Boolean)
    m_BackAutoColor = New_BackAutoColor
    PropertyChanged "BackAutoColor"
    'Code missing
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color of an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    DrawCorner
    DrawBackground
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get BackColor2() As OLE_COLOR
Attribute BackColor2.VB_Description = "Returns/sets the background color #2 of an object."
Attribute BackColor2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor2 = m_BackColor2
End Property
Public Property Let BackColor2(ByVal New_BackColor2 As OLE_COLOR)
    m_BackColor2 = New_BackColor2
    PropertyChanged "BackColor2"
    DrawCorner
    DrawBackground
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get BackColor3() As OLE_COLOR
Attribute BackColor3.VB_Description = "Returns/sets the background color #3 of an object."
Attribute BackColor3.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor3 = m_BackColor3
End Property
Public Property Let BackColor3(ByVal New_BackColor3 As OLE_COLOR)
    m_BackColor3 = New_BackColor3
    PropertyChanged "BackColor3"
    DrawCorner
    DrawBackground
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get BackColor4() As OLE_COLOR
Attribute BackColor4.VB_Description = "Returns/sets the background color #4 of an object."
Attribute BackColor4.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor4 = m_BackColor4
End Property
Public Property Let BackColor4(ByVal New_BackColor4 As OLE_COLOR)
    m_BackColor4 = New_BackColor4
    PropertyChanged "BackColor4"
    DrawCorner
    DrawBackground
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get BackColors() As mpbdColors
Attribute BackColors.VB_Description = "Returns/sets how many background colors are used in an object."
Attribute BackColors.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColors = m_BackColors
End Property
Public Property Let BackColors(ByVal New_BackColors As mpbdColors)
    m_BackColors = New_BackColors
    PropertyChanged "BackColors"
    DrawCorner
    DrawBackground
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get BarAutoColor() As Boolean
Attribute BarAutoColor.VB_Description = "Returns/sets the setting used to determine if bar and barshadow colors are automatically calculated."
Attribute BarAutoColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarAutoColor = m_BarAutoColor
End Property
Public Property Let BarAutoColor(ByVal New_BarAutoColor As Boolean)
    m_BarAutoColor = New_BarAutoColor
    PropertyChanged "BarAutoColor"
    'Code missing
End Property
Public Property Get BarColor() As OLE_COLOR
Attribute BarColor.VB_Description = "Returns/sets the progress bar color of an object."
Attribute BarColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarColor = m_BarColor
End Property
Public Property Let BarColor(ByVal New_BarColor As OLE_COLOR)
    m_BarColor = New_BarColor
    PropertyChanged "BarColor"
    DrawBarAndShadow
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get BarColor2() As OLE_COLOR
Attribute BarColor2.VB_Description = "Returns/sets the progress bar color #2 of an object."
Attribute BarColor2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarColor2 = m_BarColor2
End Property
Public Property Let BarColor2(ByVal New_BarColor2 As OLE_COLOR)
    m_BarColor2 = New_BarColor2
    PropertyChanged "BarColor2"
    DrawBarAndShadow
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get BarColor3() As OLE_COLOR
Attribute BarColor3.VB_Description = "Returns/sets the progress bar color #3 of an object."
Attribute BarColor3.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarColor3 = m_BarColor3
End Property
Public Property Let BarColor3(ByVal New_BarColor3 As OLE_COLOR)
    m_BarColor3 = New_BarColor3
    PropertyChanged "BarColor3"
    DrawBarAndShadow
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get BarColor4() As OLE_COLOR
Attribute BarColor4.VB_Description = "Returns/sets the progress bar color #4 of an object."
Attribute BarColor4.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarColor4 = m_BarColor4
End Property
Public Property Let BarColor4(ByVal New_BarColor4 As OLE_COLOR)
    m_BarColor4 = New_BarColor4
    PropertyChanged "BarColor4"
    DrawBarAndShadow
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get BarColors() As mpbdColors
Attribute BarColors.VB_Description = "Returns/sets how many progress bar colors are used in an object."
Attribute BarColors.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarColors = m_BarColors
End Property
Public Property Let BarColors(ByVal New_BarColors As mpbdColors)
    m_BarColors = New_BarColors
    PropertyChanged "BarColors"
    DrawBarAndShadow
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get BarStyle() As mpbdBarStyle
Attribute BarStyle.VB_Description = "Returns/sets style of a progress bar of an object."
Attribute BarStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarStyle = m_BarStyle
End Property
Public Property Let BarStyle(ByVal New_BarStyle As mpbdBarStyle)
    m_BarStyle = New_BarStyle
    PropertyChanged "BarStyle"
    DrawBarAndShadow
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get BarStyleFullOnly() As Boolean
Attribute BarStyleFullOnly.VB_Description = "Returns/sets if progress bar increases continuosly or by ""box by box""."
Attribute BarStyleFullOnly.VB_ProcData.VB_Invoke_Property = ";Behavior"
    BarStyle = m_BarStyle
End Property
Public Property Let BarStyleFullOnly(ByVal New_BarStyleFullOnly As Boolean)
    m_BarStyleFullOnly = New_BarStyleFullOnly
    PropertyChanged "BarStyleFullOnly"
    'Code missing
End Property
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets a border color of an object."
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    DrawCorner
    DrawBackground
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get BorderStyle() As mpbdBorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets border style of an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As mpbdBorderStyle)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    DrawCorner
    DrawBackground
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get BorderWidth() As mpbdBorderWidth
Attribute BorderWidth.VB_Description = "Returns/sets border width of an object."
Attribute BorderWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderWidth = m_BorderWidth
End Property
Public Property Let BorderWidth(ByVal New_BorderWidth As mpbdBorderWidth)
    m_BorderWidth = New_BorderWidth
    PropertyChanged "BorderWidth"
    UserControl_Resize
End Property
Public Property Get Custom() As Boolean
Attribute Custom.VB_Description = "Returns/sets if a custom text is to be displayed when finished."
Attribute Custom.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Custom = m_Custom
End Property
Public Property Let Custom(ByVal New_Custom As Boolean)
    m_Custom = New_Custom
    PropertyChanged "Custom"
    TextRefresh
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get CustomText() As String
Attribute CustomText.VB_Description = "Returns/sets the custom text to be displayed when finished."
Attribute CustomText.VB_ProcData.VB_Invoke_Property = ";Behavior"
    CustomText = m_CustomText
End Property
Public Property Let CustomText(ByVal New_CustomText As String)
    m_CustomText = New_CustomText
    PropertyChanged "CustomText"
    TextRefresh
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get Decimals() As Byte
Attribute Decimals.VB_Description = "Returns/sets number of decimals shown in percentage value."
Attribute Decimals.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Decimals = m_Decimals
End Property
Public Property Let Decimals(ByVal New_Decimals As Byte)
    m_Decimals = New_Decimals
    PropertyChanged "Decimals"
    TextRefresh
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get Direction() As mpbdDirection
Attribute Direction.VB_Description = "Returns/sets the direction progress bar increases."
Attribute Direction.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Direction = m_Direction
End Property
Public Property Let Direction(ByVal New_Direction As mpbdDirection)
    m_Direction = New_Direction
    PropertyChanged "Direction"
    Select Case m_Direction
        Case 0, 3
            FG.ForeColor = m_PercentColorBar
            SG.ForeColor = m_PercentColorShadow
        Case 1, 2
            SG.ForeColor = m_PercentColorBar
            FG.ForeColor = m_PercentColorShadow
    End Select
    Refresh
    If Not m_ManualRefresh Then Draw
End Property
Private Sub Draw()
    'clear FG and SG
    Select Case m_Direction
        Case 0 'right
            Call BitBlt(FG.hdc, 1, 0, SG.ScaleWidth, SG.ScaleHeight, Bar.hdc, 0, 0, vbSrcCopy)
            Call BitBlt(SG.hdc, 0, 0, SG.ScaleWidth, SG.ScaleHeight, Shadow.hdc, 0, 0, vbSrcCopy)
        Case 1 'up
            Call BitBlt(FG.hdc, 0, 1, SG.ScaleWidth, SG.ScaleHeight, Shadow.hdc, 0, 0, vbSrcCopy)
            Call BitBlt(SG.hdc, 0, 0, SG.ScaleWidth, SG.ScaleHeight, Bar.hdc, 0, 0, vbSrcCopy)
        Case 2 'left
            Call BitBlt(FG.hdc, 1, 0, SG.ScaleWidth, SG.ScaleHeight, Shadow.hdc, 0, 0, vbSrcCopy)
            Call BitBlt(SG.hdc, 0, 0, SG.ScaleWidth, SG.ScaleHeight, Bar.hdc, 0, 0, vbSrcCopy)
        Case 3 'down
            Call BitBlt(FG.hdc, 0, 1, SG.ScaleWidth, SG.ScaleHeight, Bar.hdc, 0, 0, vbSrcCopy)
            Call BitBlt(SG.hdc, 0, 0, SG.ScaleWidth, SG.ScaleHeight, Shadow.hdc, 0, 0, vbSrcCopy)
    End Select
    'update text
    TextRefresh
    'uh, this complicated thing takes care the text is placed correctly, no matter what is the direction and text alignment
    If m_Direction = 0 Or m_Direction = 2 Then
        SG.CurrentY = (SG.ScaleHeight - SG.TextHeight(Text)) / 2
        FG.CurrentY = SG.CurrentY
        Select Case m_PercentAlign
            Case 1
                If m_Direction = 0 Then
                    SG.CurrentX = FG.Width
                    SG.Print Text
                Else
                    FG.CurrentX = FG.ScaleWidth - FG.TextWidth(Text)
                    FG.Print Text
                End If
            Case 2
                If m_Direction = 0 Then
                    FG.CurrentX = FG.ScaleWidth - FG.TextWidth(Text)
                    FG.Print Text
                Else
                    SG.CurrentX = FG.Width
                    SG.Print Text
                End If
            Case 3
                If m_Direction = 0 Then
                    SG.CurrentX = FG.Width + ((SG.ScaleWidth - FG.Width) - SG.TextWidth(Text)) / 2
                    SG.Print Text
                Else
                    FG.CurrentX = (FG.ScaleWidth - FG.TextWidth(Text)) / 2
                    FG.Print Text
                End If
            Case 4
                If m_Direction = 0 Then
                    FG.CurrentX = (FG.ScaleWidth - FG.TextWidth(Text)) / 2
                    FG.Print Text
                Else
                    SG.CurrentX = FG.Width + ((SG.ScaleWidth - FG.Width) - SG.TextWidth(Text)) / 2
                    SG.Print Text
                End If
            Case 5
                SG.CurrentX = m_BarColors + 1
                FG.CurrentX = SG.CurrentX - FG.Left
                SG.Print Text
                FG.Print Text
            Case 6
                SG.CurrentX = SG.ScaleWidth - SG.TextWidth(Text) - m_BarColors - 1
                FG.CurrentX = SG.CurrentX - FG.Left
                SG.Print Text
                FG.Print Text
            Case Else
                SG.CurrentX = (SG.ScaleWidth - SG.TextWidth(Text)) / 2
                FG.CurrentX = SG.CurrentX - FG.Left
                SG.Print Text
                FG.Print Text
        End Select
    Else
        SG.CurrentX = (SG.ScaleWidth - SG.TextWidth(Text)) / 2
        FG.CurrentX = SG.CurrentX
        Select Case m_PercentAlign
            Case 1
                If m_Direction = 3 Then
                    SG.CurrentY = FG.Height
                    SG.Print Text
                Else
                    FG.CurrentY = FG.ScaleHeight - FG.TextHeight(Text)
                    FG.Print Text
                End If
            Case 2
                If m_Direction = 3 Then
                    FG.CurrentY = FG.ScaleHeight - FG.TextHeight(Text)
                    FG.Print Text
                Else
                    SG.CurrentY = FG.Height
                    SG.Print Text
                End If
            Case 3
                If m_Direction = 3 Then
                    SG.CurrentY = FG.Height + ((SG.ScaleHeight - FG.Height) - SG.TextHeight(Text)) / 2
                    SG.Print Text
                Else
                    FG.CurrentY = (FG.ScaleHeight - FG.TextHeight(Text)) / 2
                    FG.Print Text
                End If
            Case 4
                If m_Direction = 3 Then
                    FG.CurrentY = (FG.ScaleHeight - FG.TextHeight(Text)) / 2
                    FG.Print Text
                Else
                    SG.CurrentY = FG.Height + ((SG.ScaleHeight - FG.Height) - SG.TextHeight(Text)) / 2
                    SG.Print Text
                End If
            Case 5
                SG.CurrentY = m_BarColors + 1
                FG.CurrentY = SG.CurrentY - FG.Top
                SG.Print Text
                FG.Print Text
            Case 6
                SG.CurrentY = SG.ScaleHeight - SG.TextHeight(Text) - m_BarColors - 1
                FG.CurrentY = SG.CurrentY - FG.Top
                SG.Print Text
                FG.Print Text
            Case Else
                SG.CurrentY = (SG.ScaleHeight - SG.TextHeight(Text)) / 2
                FG.CurrentY = SG.CurrentY - FG.Top
                SG.Print Text
                FG.Print Text
        End Select
    End If
    SG.Refresh
    FG.Refresh
End Sub
Public Sub DrawBackground()
    Dim TempX As Byte, TempY As Byte
    'set backcolor
    BG.BackColor = m_BackColor
    'in case of too small size, make sure borders are painted correctly - otherwise choose maximum size
    If BG.ScaleWidth < 16 Then TempX = BG.ScaleWidth / 2 Else TempX = 8
    If BG.ScaleHeight < 16 Then TempY = BG.ScaleWidth / 2 Else TempY = 8
    'draw the corners
    Call BitBlt(BG.hdc, 0, 0, TempX, TempY, Corner.hdc, 0, 0, vbSrcCopy)
    Call BitBlt(BG.hdc, BG.ScaleWidth - TempX, 0, TempX, TempY, Corner.hdc, 16 - TempX, 0, vbSrcCopy)
    Call BitBlt(BG.hdc, 0, BG.ScaleHeight - TempY, TempX, TempY, Corner.hdc, 0, 16 - TempY, vbSrcCopy)
    Call BitBlt(BG.hdc, BG.ScaleWidth - TempX, BG.ScaleHeight - TempY, TempX, TempY, Corner.hdc, 16 - TempX, 16 - TempY, vbSrcCopy)
    'draw the borders
    If BG.ScaleWidth > 16 Then
        Call StretchBlt(BG.hdc, 8, 0, BG.ScaleWidth - 16, TempY, Corner.hdc, 7, 0, 1, TempY, vbSrcCopy)
        Call StretchBlt(BG.hdc, 8, BG.ScaleHeight - TempY, BG.ScaleWidth - 16, TempY, Corner.hdc, 7, 16 - TempY, 1, TempY, vbSrcCopy)
    End If
    If BG.ScaleHeight > 16 Then
        Call StretchBlt(BG.hdc, 0, 8, TempX, BG.ScaleHeight - 16, Corner.hdc, 0, 7, TempX, 1, vbSrcCopy)
        Call StretchBlt(BG.hdc, BG.ScaleWidth - TempY, 8, TempX, BG.ScaleHeight - 16, Corner.hdc, 16 - TempX, 7, TempX, 1, vbSrcCopy)
    End If
    BG.Refresh
End Sub
Private Sub DrawBarAndShadow()
    Dim Temp As Byte
    Bar.BackColor = m_BarColor
    Shadow.BackColor = m_ShadowColor
    'draw the borders
    If m_BarColors > 0 Then
        Temp = m_BarColors - 1
        Call SetPixel(Bar.hdc, Temp, Temp, m_BarColor2)
        Call StretchBlt(Bar.hdc, 1 + Temp, Temp, Bar.ScaleWidth - 1 - Temp * 2, 1, Bar.hdc, Temp, Temp, 1, 1, vbSrcCopy)
        Call StretchBlt(Bar.hdc, Temp, 1 + Temp, 1, Bar.ScaleHeight - 1 - Temp * 2, Bar.hdc, Temp, Temp, 1, 1, vbSrcCopy)
        Call StretchBlt(Bar.hdc, 1 + Temp, Bar.ScaleHeight - 1 - Temp, Bar.ScaleWidth - 1 - Temp * 2, 1, Bar.hdc, Temp, Temp, 1, 1, vbSrcCopy)
        Call StretchBlt(Bar.hdc, Bar.ScaleWidth - 1 - Temp, 1 + Temp, 1, Bar.ScaleHeight - 1 - Temp * 2, Bar.hdc, Temp, Temp, 1, 1, vbSrcCopy)
        Call SetPixel(Shadow.hdc, Temp, Temp, m_ShadowColor2)
        Call StretchBlt(Shadow.hdc, 1 + Temp, Temp, Shadow.ScaleWidth - 1 - Temp * 2, 1, Shadow.hdc, Temp, Temp, 1, 1, vbSrcCopy)
        Call StretchBlt(Shadow.hdc, Temp, 1 + Temp, 1, Shadow.ScaleHeight - 1 - Temp * 2, Shadow.hdc, Temp, Temp, 1, 1, vbSrcCopy)
        Call StretchBlt(Shadow.hdc, 1 + Temp, Shadow.ScaleHeight - 1 - Temp, Shadow.ScaleWidth - 1 - Temp * 2, 1, Shadow.hdc, Temp, Temp, 1, 1, vbSrcCopy)
        Call StretchBlt(Shadow.hdc, Shadow.ScaleWidth - 1 - Temp, 1 + Temp, 1, Shadow.ScaleHeight - 1 - Temp * 2, Shadow.hdc, Temp, Temp, 1, 1, vbSrcCopy)
    End If
    If m_BarColors > 1 Then
        Temp = m_BarColors - 2
        Call SetPixel(Bar.hdc, Temp, Temp, m_BarColor3)
        Call StretchBlt(Bar.hdc, 1 + Temp, Temp, Bar.ScaleWidth - 1 - Temp * 2, 1, Bar.hdc, Temp, Temp, 1, 1, vbSrcCopy)
        Call StretchBlt(Bar.hdc, Temp, 1 + Temp, 1, Bar.ScaleHeight - 1 - Temp * 2, Bar.hdc, Temp, Temp, 1, 1, vbSrcCopy)
        Call StretchBlt(Bar.hdc, 1 + Temp, Bar.ScaleHeight - 1 - Temp, Bar.ScaleWidth - 1 - Temp * 2, 1, Bar.hdc, Temp, Temp, 1, 1, vbSrcCopy)
        Call StretchBlt(Bar.hdc, Bar.ScaleWidth - 1 - Temp, 1 + Temp, 1, Bar.ScaleHeight - 1 - Temp * 2, Bar.hdc, Temp, Temp, 1, 1, vbSrcCopy)
        Call SetPixel(Shadow.hdc, Temp, Temp, m_ShadowColor3)
        Call StretchBlt(Shadow.hdc, 1 + Temp, Temp, Shadow.ScaleWidth - 1 - Temp * 2, 1, Shadow.hdc, Temp, Temp, 1, 1, vbSrcCopy)
        Call StretchBlt(Shadow.hdc, Temp, 1 + Temp, 1, Shadow.ScaleHeight - 1 - Temp * 2, Shadow.hdc, Temp, Temp, 1, 1, vbSrcCopy)
        Call StretchBlt(Shadow.hdc, 1 + Temp, Shadow.ScaleHeight - 1 - Temp, Shadow.ScaleWidth - 1 - Temp * 2, 1, Shadow.hdc, Temp, Temp, 1, 1, vbSrcCopy)
        Call StretchBlt(Shadow.hdc, Shadow.ScaleWidth - 1 - Temp, 1 + Temp, 1, Shadow.ScaleHeight - 1 - Temp * 2, Shadow.hdc, Temp, Temp, 1, 1, vbSrcCopy)
    End If
    If m_BarColors > 2 Then
        Call SetPixel(Bar.hdc, 0, 0, m_BarColor4)
        Call StretchBlt(Bar.hdc, 1, 0, Bar.ScaleWidth - 1, 1, Bar.hdc, 0, 0, 1, 1, vbSrcCopy)
        Call StretchBlt(Bar.hdc, 0, 1, 1, Bar.ScaleHeight - 1, Bar.hdc, 0, 0, 1, 1, vbSrcCopy)
        Call StretchBlt(Bar.hdc, 1, Bar.ScaleHeight - 1, Bar.ScaleWidth - 1, 1, Bar.hdc, 0, 0, 1, 1, vbSrcCopy)
        Call StretchBlt(Bar.hdc, Bar.ScaleWidth - 1, 1, 1, Bar.ScaleHeight - 1, Bar.hdc, 0, 0, 1, 1, vbSrcCopy)
        Call SetPixel(Shadow.hdc, 0, 0, m_ShadowColor4)
        Call StretchBlt(Shadow.hdc, 1, 0, Shadow.ScaleWidth - 1, 1, Shadow.hdc, 0, 0, 1, 1, vbSrcCopy)
        Call StretchBlt(Shadow.hdc, 0, 1, 1, Shadow.ScaleHeight - 1, Shadow.hdc, 0, 0, 1, 1, vbSrcCopy)
        Call StretchBlt(Shadow.hdc, 1, Shadow.ScaleHeight - 1, Shadow.ScaleWidth - 1, 1, Shadow.hdc, 0, 0, 1, 1, vbSrcCopy)
        Call StretchBlt(Shadow.hdc, Shadow.ScaleWidth - 1, 1, 1, Shadow.ScaleHeight - 1, Shadow.hdc, 0, 0, 1, 1, vbSrcCopy)
    End If
End Sub
Private Sub DrawCorner()
    Dim TempColor As Long
    Select Case m_BorderStyle
        Case 0 'draw normal border
            Corner.BackColor = m_BackColor
            Call SetPixel(Corner.hdc, 0, 0, m_BorderColor)
            Call StretchBlt(Corner.hdc, 1, 0, Corner.ScaleWidth - 1, 1, Corner.hdc, 0, 0, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, 1, Corner.ScaleHeight - 1, Corner.ScaleWidth - 1, 1, Corner.hdc, 0, 0, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, 0, 1, 1, Corner.ScaleHeight - 1, Corner.hdc, 0, 0, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 1, 1, 1, Corner.ScaleHeight - 1, Corner.hdc, 0, 0, 1, 1, vbSrcCopy)
            Select Case m_BackColors
                Case 1
                    Call SetPixel(Corner.hdc, 1, 1, m_BackColor2)
                    Call StretchBlt(Corner.hdc, 2, 1, Corner.ScaleWidth - 3, 1, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 2, Corner.ScaleHeight - 2, Corner.ScaleWidth - 3, 1, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 1, 2, 1, Corner.ScaleHeight - 3, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 2, 2, 1, Corner.ScaleHeight - 3, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
                Case 2
                    Call SetPixel(Corner.hdc, 1, 1, m_BackColor3)
                    Call StretchBlt(Corner.hdc, 2, 1, Corner.ScaleWidth - 3, 1, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 2, Corner.ScaleHeight - 2, Corner.ScaleWidth - 3, 1, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 1, 2, 1, Corner.ScaleHeight - 3, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 2, 2, 1, Corner.ScaleHeight - 3, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
                    
                    Call SetPixel(Corner.hdc, 2, 2, m_BackColor2)
                    Call StretchBlt(Corner.hdc, 3, 2, Corner.ScaleWidth - 5, 1, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 3, Corner.ScaleHeight - 3, Corner.ScaleWidth - 5, 1, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 2, 3, 1, Corner.ScaleHeight - 5, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 3, 3, 1, Corner.ScaleHeight - 5, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                Case 3
                    Call SetPixel(Corner.hdc, 1, 1, m_BackColor4)
                    Call StretchBlt(Corner.hdc, 2, 1, Corner.ScaleWidth - 3, 1, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 2, Corner.ScaleHeight - 2, Corner.ScaleWidth - 3, 1, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 1, 2, 1, Corner.ScaleHeight - 3, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 2, 2, 1, Corner.ScaleHeight - 3, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
                    
                    Call SetPixel(Corner.hdc, 2, 2, m_BackColor3)
                    Call StretchBlt(Corner.hdc, 3, 2, Corner.ScaleWidth - 5, 1, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 3, Corner.ScaleHeight - 3, Corner.ScaleWidth - 5, 1, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 2, 3, 1, Corner.ScaleHeight - 5, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 3, 3, 1, Corner.ScaleHeight - 5, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    
                    Call SetPixel(Corner.hdc, 3, 3, m_BackColor2)
                    Call StretchBlt(Corner.hdc, 4, 3, Corner.ScaleWidth - 7, 1, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 4, Corner.ScaleHeight - 4, Corner.ScaleWidth - 7, 1, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 3, 4, 1, Corner.ScaleHeight - 7, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 4, 4, 1, Corner.ScaleHeight - 7, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
            End Select
        Case 1 'draw slighty rounded border
            'get windows' backcolor
            Corner.BackColor = m_FormColor
            TempColor = GetPixel(Corner.hdc, 0, 0)
            'set object backcolor
            Corner.BackColor = m_BackColor
            'draw form backcolor to very corners
            Call SetPixel(Corner.hdc, 0, 0, TempColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 1, 0, TempColor)
            Call SetPixel(Corner.hdc, 0, Corner.ScaleHeight - 1, TempColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 1, Corner.ScaleHeight - 1, TempColor)
            Call SetPixel(Corner.hdc, 1, 0, TempColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 2, 0, TempColor)
            Call SetPixel(Corner.hdc, 1, Corner.ScaleHeight - 1, TempColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 2, Corner.ScaleHeight - 1, TempColor)
            Call SetPixel(Corner.hdc, 0, 1, TempColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 1, 1, TempColor)
            Call SetPixel(Corner.hdc, 0, Corner.ScaleHeight - 2, TempColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 1, Corner.ScaleHeight - 2, TempColor)
            'corners
            Call SetPixel(Corner.hdc, 1, 1, m_BorderColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 2, 1, m_BorderColor)
            Call SetPixel(Corner.hdc, 1, Corner.ScaleHeight - 2, m_BorderColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 2, Corner.ScaleHeight - 2, m_BorderColor)
            'borders
            Call StretchBlt(Corner.hdc, 2, 0, Corner.ScaleWidth - 4, 1, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, 2, Corner.ScaleHeight - 1, Corner.ScaleWidth - 4, 1, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, 0, 2, 1, Corner.ScaleHeight - 4, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 1, 2, 1, Corner.ScaleHeight - 4, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
            Select Case m_BackColors
                Case 1
                    Call SetPixel(Corner.hdc, 2, 1, m_BackColor2)
                    Call StretchBlt(Corner.hdc, 3, 1, Corner.ScaleWidth - 5, 1, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 1, 2, 1, Corner.ScaleHeight - 4, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 2, Corner.ScaleHeight - 2, Corner.ScaleWidth - 4, 1, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 2, 2, 1, Corner.ScaleHeight - 4, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
                Case 2
                    Call SetPixel(Corner.hdc, 2, 1, m_BackColor3)
                    Call StretchBlt(Corner.hdc, 3, 1, Corner.ScaleWidth - 5, 1, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 1, 2, 1, Corner.ScaleHeight - 4, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 2, Corner.ScaleHeight - 2, Corner.ScaleWidth - 4, 1, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 2, 2, 1, Corner.ScaleHeight - 4, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
                    
                    Call SetPixel(Corner.hdc, 2, 2, m_BackColor2)
                    Call StretchBlt(Corner.hdc, 3, 2, Corner.ScaleWidth - 5, 1, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 3, Corner.ScaleHeight - 3, Corner.ScaleWidth - 5, 1, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 2, 3, 1, Corner.ScaleHeight - 5, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 3, 3, 1, Corner.ScaleHeight - 5, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                Case 3
                    Call SetPixel(Corner.hdc, 2, 1, m_BackColor4)
                    Call StretchBlt(Corner.hdc, 3, 1, Corner.ScaleWidth - 5, 1, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 1, 2, 1, Corner.ScaleHeight - 4, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 2, Corner.ScaleHeight - 2, Corner.ScaleWidth - 4, 1, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 2, 2, 1, Corner.ScaleHeight - 4, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
                    
                    Call SetPixel(Corner.hdc, 2, 2, m_BackColor3)
                    Call StretchBlt(Corner.hdc, 3, 2, Corner.ScaleWidth - 5, 1, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 3, Corner.ScaleHeight - 3, Corner.ScaleWidth - 5, 1, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 2, 3, 1, Corner.ScaleHeight - 5, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 3, 3, 1, Corner.ScaleHeight - 5, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    
                    Call SetPixel(Corner.hdc, 3, 3, m_BackColor2)
                    Call StretchBlt(Corner.hdc, 4, 3, Corner.ScaleWidth - 7, 1, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 4, Corner.ScaleHeight - 4, Corner.ScaleWidth - 7, 1, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 3, 4, 1, Corner.ScaleHeight - 7, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 4, 4, 1, Corner.ScaleHeight - 7, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
            End Select
        Case 2 'draw rounded border
            'get windows' backcolor
            Corner.BackColor = m_FormColor
            TempColor = GetPixel(Corner.hdc, 0, 0)
            'set backcolor
            Corner.BackColor = m_BackColor
            'draw form backcolor to very corners
            Call SetPixel(Corner.hdc, 1, 1, TempColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 2, 1, TempColor)
            Call SetPixel(Corner.hdc, 1, Corner.ScaleHeight - 2, TempColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 2, Corner.ScaleHeight - 2, TempColor)
            Call StretchBlt(Corner.hdc, 0, 0, 4, 1, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, 0, 1, 1, 3, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 4, 0, 4, 1, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 1, 1, 1, 3, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, 0, Corner.ScaleHeight - 1, 4, 1, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, 0, Corner.ScaleHeight - 4, 1, 3, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 4, Corner.ScaleHeight - 1, 4, 1, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 1, Corner.ScaleHeight - 4, 1, 3, Corner.hdc, 1, 1, 1, 1, vbSrcCopy)
            'corner #1
            Call SetPixel(Corner.hdc, 2, 1, m_BorderColor)
            Call SetPixel(Corner.hdc, 3, 1, m_BorderColor)
            Call SetPixel(Corner.hdc, 1, 2, m_BorderColor)
            Call SetPixel(Corner.hdc, 1, 3, m_BorderColor)
            'corner #2
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 3, 1, m_BorderColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 4, 1, m_BorderColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 2, 2, m_BorderColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 2, 3, m_BorderColor)
            'corner #3
            Call SetPixel(Corner.hdc, 2, Corner.ScaleHeight - 2, m_BorderColor)
            Call SetPixel(Corner.hdc, 3, Corner.ScaleHeight - 2, m_BorderColor)
            Call SetPixel(Corner.hdc, 1, Corner.ScaleHeight - 3, m_BorderColor)
            Call SetPixel(Corner.hdc, 1, Corner.ScaleHeight - 4, m_BorderColor)
            'corner #4
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 3, Corner.ScaleHeight - 2, m_BorderColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 4, Corner.ScaleHeight - 2, m_BorderColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 2, Corner.ScaleHeight - 3, m_BorderColor)
            Call SetPixel(Corner.hdc, Corner.ScaleWidth - 2, Corner.ScaleHeight - 4, m_BorderColor)
            'borders
            Call StretchBlt(Corner.hdc, 4, 0, Corner.ScaleWidth - 8, 1, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, 4, Corner.ScaleHeight - 1, Corner.ScaleWidth - 8, 1, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, 0, 4, 1, Corner.ScaleHeight - 8, Corner.hdc, 2, 1, 1, 1, vbSrcCopy)
            Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 1, 4, 1, Corner.ScaleHeight - 8, Corner.hdc, 2, 1, 1, 1, vbSrcCopy) '
            Select Case m_BackColors
                Case 1
                    Call SetPixel(Corner.hdc, 2, 2, m_BackColor2)
                    Call SetPixel(Corner.hdc, 2, 3, m_BackColor2)
                    Call SetPixel(Corner.hdc, 3, 2, m_BackColor2)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 3, 2, m_BackColor2)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 3, 3, m_BackColor2)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 4, 2, m_BackColor2)
                    Call SetPixel(Corner.hdc, 2, Corner.ScaleHeight - 3, m_BackColor2)
                    Call SetPixel(Corner.hdc, 2, Corner.ScaleHeight - 4, m_BackColor2)
                    Call SetPixel(Corner.hdc, 3, Corner.ScaleHeight - 3, m_BackColor2)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 3, Corner.ScaleHeight - 3, m_BackColor2)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 3, Corner.ScaleHeight - 4, m_BackColor2)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 4, Corner.ScaleHeight - 3, m_BackColor2)
                    Call StretchBlt(Corner.hdc, 4, 1, Corner.ScaleWidth - 8, 1, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 4, Corner.ScaleHeight - 2, Corner.ScaleWidth - 8, 1, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 1, 4, 1, Corner.ScaleHeight - 8, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 2, 4, 1, Corner.ScaleHeight - 8, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                Case 2
                    Call SetPixel(Corner.hdc, 2, 2, m_BackColor3)
                    Call SetPixel(Corner.hdc, 2, 3, m_BackColor3)
                    Call SetPixel(Corner.hdc, 3, 2, m_BackColor3)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 3, 2, m_BackColor3)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 3, 3, m_BackColor3)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 4, 2, m_BackColor3)
                    Call SetPixel(Corner.hdc, 2, Corner.ScaleHeight - 3, m_BackColor3)
                    Call SetPixel(Corner.hdc, 2, Corner.ScaleHeight - 4, m_BackColor3)
                    Call SetPixel(Corner.hdc, 3, Corner.ScaleHeight - 3, m_BackColor3)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 3, Corner.ScaleHeight - 3, m_BackColor3)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 3, Corner.ScaleHeight - 4, m_BackColor3)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 4, Corner.ScaleHeight - 3, m_BackColor3)
                    Call StretchBlt(Corner.hdc, 4, 1, Corner.ScaleWidth - 8, 1, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 4, Corner.ScaleHeight - 2, Corner.ScaleWidth - 8, 1, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 1, 4, 1, Corner.ScaleHeight - 8, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 2, 4, 1, Corner.ScaleHeight - 8, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    
                    Call SetPixel(Corner.hdc, 3, 3, m_BackColor2)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 4, 3, m_BackColor2)
                    Call SetPixel(Corner.hdc, 3, Corner.ScaleHeight - 4, m_BackColor2)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 4, Corner.ScaleHeight - 4, m_BackColor2)
                    Call StretchBlt(Corner.hdc, 4, 2, Corner.ScaleWidth - 8, 1, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 4, Corner.ScaleHeight - 3, Corner.ScaleWidth - 8, 1, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 2, 4, 1, Corner.ScaleHeight - 8, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 3, 4, 1, Corner.ScaleHeight - 8, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
                Case 3
                    Call SetPixel(Corner.hdc, 2, 2, m_BackColor4)
                    Call SetPixel(Corner.hdc, 2, 3, m_BackColor4)
                    Call SetPixel(Corner.hdc, 3, 2, m_BackColor4)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 3, 2, m_BackColor4)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 3, 3, m_BackColor4)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 4, 2, m_BackColor4)
                    Call SetPixel(Corner.hdc, 2, Corner.ScaleHeight - 3, m_BackColor4)
                    Call SetPixel(Corner.hdc, 2, Corner.ScaleHeight - 4, m_BackColor4)
                    Call SetPixel(Corner.hdc, 3, Corner.ScaleHeight - 3, m_BackColor4)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 3, Corner.ScaleHeight - 3, m_BackColor4)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 3, Corner.ScaleHeight - 4, m_BackColor4)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 4, Corner.ScaleHeight - 3, m_BackColor4)
                    Call StretchBlt(Corner.hdc, 4, 1, Corner.ScaleWidth - 8, 1, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 4, Corner.ScaleHeight - 2, Corner.ScaleWidth - 8, 1, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 1, 4, 1, Corner.ScaleHeight - 8, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 2, 4, 1, Corner.ScaleHeight - 8, Corner.hdc, 2, 2, 1, 1, vbSrcCopy)
                    
                    Call SetPixel(Corner.hdc, 3, 3, m_BackColor3)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 4, 3, m_BackColor3)
                    Call SetPixel(Corner.hdc, 3, Corner.ScaleHeight - 4, m_BackColor3)
                    Call SetPixel(Corner.hdc, Corner.ScaleWidth - 4, Corner.ScaleHeight - 4, m_BackColor3)
                    Call StretchBlt(Corner.hdc, 4, 2, Corner.ScaleWidth - 8, 1, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 4, Corner.ScaleHeight - 3, Corner.ScaleWidth - 8, 1, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 2, 4, 1, Corner.ScaleHeight - 8, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 3, 4, 1, Corner.ScaleHeight - 8, Corner.hdc, 3, 3, 1, 1, vbSrcCopy)
                    
                    Call SetPixel(Corner.hdc, 4, 3, m_BackColor2)
                    Call StretchBlt(Corner.hdc, 5, 3, Corner.ScaleWidth - 9, 1, Corner.hdc, 4, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 3, 4, 1, Corner.ScaleHeight - 8, Corner.hdc, 4, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, 4, Corner.ScaleHeight - 4, Corner.ScaleWidth - 8, 1, Corner.hdc, 4, 3, 1, 1, vbSrcCopy)
                    Call StretchBlt(Corner.hdc, Corner.ScaleWidth - 4, 4, 1, Corner.ScaleHeight - 8, Corner.hdc, 4, 3, 1, 1, vbSrcCopy)
            End Select
    End Select
End Sub
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns/sets the font of an object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
    Set Font = m_Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    Set SG.Font = New_Font
    Set FG.Font = New_Font
    'make sure font colors are correct
    Select Case m_Direction
        Case 0, 3
            FG.ForeColor = m_PercentColorBar
            SG.ForeColor = m_PercentColorShadow
        Case 1, 2
            SG.ForeColor = m_PercentColorBar
            FG.ForeColor = m_PercentColorShadow
    End Select
    PropertyChanged "Font"
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get FormColor() As OLE_COLOR
Attribute FormColor.VB_Description = "Returns/sets the real background of an object."
Attribute FormColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FormColor = m_FormColor
End Property
Public Property Let FormColor(ByVal New_FormColor As OLE_COLOR)
    m_FormColor = New_FormColor
    PropertyChanged "FormColor"
    DrawCorner
    DrawBackground
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get ManualRefresh() As Boolean
Attribute ManualRefresh.VB_Description = "Returns/sets if object is manually refreshed."
Attribute ManualRefresh.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ManualRefresh = m_ManualRefresh
End Property
Public Property Let ManualRefresh(ByVal New_ManualRefresh As Boolean)
    m_ManualRefresh = New_ManualRefresh
    PropertyChanged "ManualRefresh"
End Property
Public Property Get Max() As Currency
Attribute Max.VB_Description = "Returns/sets the maximum value of an object."
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Misc"
    Max = m_Max
End Property
Public Property Let Max(ByVal New_Max As Currency)
    If New_Max < m_Min + 1 Then New_Max = m_Min + 1
    m_Max = New_Max
    PropertyChanged "Max"
    If m_Max < m_Value Then m_Value = m_Max
    If Not m_ManualRefresh Then Refresh: Draw
End Property
Public Property Get Min() As Currency
Attribute Min.VB_Description = "Returns/sets the minimum value of an object."
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Misc"
    Min = m_Min
End Property
Public Property Let Min(ByVal New_Min As Currency)
    If New_Min > m_Max - 1 Then New_Min = m_Max - 1
    m_Min = New_Min
    PropertyChanged "Min"
    If m_Min > m_Value Then m_Value = m_Min
    If Not m_ManualRefresh Then Refresh: Draw
End Property
Public Property Get NoPercent() As Boolean
Attribute NoPercent.VB_Description = "Returns/sets if percentage is shown."
Attribute NoPercent.VB_ProcData.VB_Invoke_Property = ";Appearance"
    NoPercent = m_NoPercent
End Property
Public Property Let NoPercent(ByVal New_NoPercent As Boolean)
    m_NoPercent = New_NoPercent
    PropertyChanged "NoPercent"
    TextRefresh
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get Percent() As Single
Attribute Percent.VB_Description = "Returns current percentage of an object."
Attribute Percent.VB_MemberFlags = "400"
    Percent = m_Percent
End Property
Public Property Let Percent(ByVal New_Percent As Single)
    If Ambient.UserMode = False Then Err.Raise 382
    If Ambient.UserMode Then Err.Raise 393
End Property
Public Property Get PercentAfter() As Integer
Attribute PercentAfter.VB_Description = "Returns/sets text to be shown after the percentage value of an object."
Attribute PercentAfter.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PercentAfter = m_PercentAfter
End Property
Public Property Let PercentAfter(ByVal New_PercentAfter As Integer)
    m_PercentAfter = New_PercentAfter
    PropertyChanged "PercentAfter"
    TextRefresh
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get PercentAlign() As mpbdPercentAlign
Attribute PercentAlign.VB_Description = "Returns/sets the percent bar text alignment of an object."
Attribute PercentAlign.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PercentAlign = m_PercentAlign
End Property
Public Property Let PercentAlign(ByVal New_PercentAlign As mpbdPercentAlign)
    m_PercentAlign = New_PercentAlign
    PropertyChanged "PercentAlign"
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get PercentBefore() As String
Attribute PercentBefore.VB_Description = "Returns/sets the text shown before the percentage value of an object."
Attribute PercentBefore.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PercentBefore = m_PercentBefore
End Property
Public Property Let PercentBefore(ByVal New_PercentBefore As String)
    m_PercentBefore = New_PercentBefore
    PropertyChanged "PercentBefore"
    TextRefresh
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get PercentColorBar() As OLE_COLOR
Attribute PercentColorBar.VB_Description = "Returns/sets the percentage text color on the progress bar of an object."
Attribute PercentColorBar.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PercentColorBar = m_PercentColorBar
End Property
Public Property Let PercentColorBar(ByVal New_PercentColorBar As OLE_COLOR)
    m_PercentColorBar = New_PercentColorBar
    PropertyChanged "PercentColorBar"
    Select Case m_Direction
        Case 0, 3
            FG.ForeColor = m_PercentColorBar
        Case 1, 2
            SG.ForeColor = m_PercentColorBar
    End Select
End Property
Public Property Get PercentColorShadow() As OLE_COLOR
Attribute PercentColorShadow.VB_Description = "Returns/sets the percentage text color on the progress bar shadow of an object."
Attribute PercentColorShadow.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PercentColorShadow = m_PercentColorShadow
End Property
Public Property Let PercentColorShadow(ByVal New_PercentColorShadow As OLE_COLOR)
    m_PercentColorShadow = New_PercentColorShadow
    Select Case m_Direction
        Case 0, 3
            SG.ForeColor = m_PercentColorShadow
        Case 1, 2
            FG.ForeColor = m_PercentColorShadow
    End Select
    PropertyChanged "PercentColorShadow"
End Property
Public Sub Refresh()
    Select Case m_Direction
        Case 0 'moving left
            FG.Move -1, 0, (m_Value - m_Min) / (m_Max - m_Min) * (SG.ScaleWidth + 1), SG.ScaleHeight
        Case 1 'moving up
            FG.Move 0, -1, SG.ScaleWidth, ((m_Max - m_Min) - (m_Value - m_Min)) / (m_Max - m_Min) * (SG.ScaleHeight + 1)
        Case 2 'moving right
            FG.Move -1, 0, ((m_Max - m_Min) - (m_Value - m_Min)) / (m_Max - m_Min) * (SG.ScaleWidth + 1), SG.ScaleHeight
        Case 3 'moving down
            FG.Move 0, -1, SG.ScaleWidth, (m_Value - m_Min) / (m_Max - m_Min) * (SG.ScaleHeight + 1)
    End Select
    If m_ManualRefresh Then Draw
End Sub
Public Property Get ScaleMode() As ScaleModeConstants
Attribute ScaleMode.VB_Description = "Returns/sets the scale mode of an object."
Attribute ScaleMode.VB_ProcData.VB_Invoke_Property = ";Scale"
    ScaleMode = m_ScaleMode
End Property
Public Property Get ScaleHeight() As ScaleModeConstants
Attribute ScaleHeight.VB_Description = "Returns the number of units for the vertical measurement of an object's interior."
Attribute ScaleHeight.VB_ProcData.VB_Invoke_Property = ";Scale"
    ScaleHeight = SG.ScaleY(m_ScaleHeight, SG.ScaleMode, m_ScaleMode)
End Property
Public Property Let ScaleMode(ByVal New_ScaleMode As ScaleModeConstants)
    m_ScaleMode = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property
Public Property Get ScaleWidth() As ScaleModeConstants
Attribute ScaleWidth.VB_Description = "Returns the number of units for the horizontal measurement of an object's interior."
Attribute ScaleWidth.VB_ProcData.VB_Invoke_Property = ";Scale"
    ScaleWidth = SG.ScaleX(m_ScaleWidth, SG.ScaleMode, m_ScaleMode)
End Property
Public Sub SetParent(ByVal hwnd As Long)
    SetWindowParent UserControl.hwnd, hwnd
End Sub
Public Property Get ShadowColor() As OLE_COLOR
Attribute ShadowColor.VB_Description = "Returns/sets the progress bar shadow color of an object."
Attribute ShadowColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ShadowColor = m_ShadowColor
End Property
Public Property Let ShadowColor(ByVal New_ShadowColor As OLE_COLOR)
    m_ShadowColor = New_ShadowColor
    PropertyChanged "ShadowColor"
    DrawBarAndShadow
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get ShadowColor2() As OLE_COLOR
Attribute ShadowColor2.VB_Description = "Returns/sets the progress bar shadow color #2 of an object."
Attribute ShadowColor2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ShadowColor2 = m_ShadowColor2
End Property
Public Property Let ShadowColor2(ByVal New_ShadowColor2 As OLE_COLOR)
    m_ShadowColor2 = New_ShadowColor2
    PropertyChanged "ShadowColor2"
    DrawBarAndShadow
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get ShadowColor3() As OLE_COLOR
Attribute ShadowColor3.VB_Description = "Returns/sets the progress bar shadow color #3 of an object."
Attribute ShadowColor3.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ShadowColor3 = m_ShadowColor3
End Property
Public Property Let ShadowColor3(ByVal New_ShadowColor3 As OLE_COLOR)
    m_ShadowColor3 = New_ShadowColor3
    PropertyChanged "ShadowColor3"
    DrawBarAndShadow
    If Not m_ManualRefresh Then Draw
End Property
Public Property Get ShadowColor4() As OLE_COLOR
Attribute ShadowColor4.VB_Description = "Returns/sets the progress bar shadow color #4 of an object."
Attribute ShadowColor4.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ShadowColor4 = m_ShadowColor4
End Property
Public Property Let ShadowColor4(ByVal New_ShadowColor4 As OLE_COLOR)
    m_ShadowColor4 = New_ShadowColor4
    PropertyChanged "ShadowColor4"
    DrawBarAndShadow
    If Not m_ManualRefresh Then Draw
End Property
Private Sub TextRefresh()
    Dim Temp1 As String, Temp2 As String
    If Not m_NoPercent Then
        If m_Custom And m_Value = m_Max Then
            Text = m_CustomText
        Else
            If m_Decimals > 0 Then
                Text = m_PercentBefore & Chr(32) & Format(m_Percent, "0." & Zeros(m_Decimals)) & Chr(32) & m_PercentAfter
            Else
                If m_PercentBefore <> "" Then Temp1 = m_PercentBefore & Chr(32)
                If m_PercentAfter <> "" Then Temp2 = Chr(32) & m_PercentAfter
                Text = Temp1 & Fix(m_Percent) & Temp2
            End If
        End If
    ElseIf m_PercentBefore <> "" Then
        Text = m_PercentBefore
    End If
    m_TextWidth = SG.TextWidth(Text)
    m_TextHeight = SG.TextHeight(Text)
End Sub
Public Property Get Value() As Currency
Attribute Value.VB_Description = "Returns/sets the value of an object."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
    Value = m_Value
End Property
Public Property Let Value(ByVal New_Value As Currency)
    On Error Resume Next
    If New_Value < m_Min Then New_Value = m_Min
    If New_Value > m_Max Then New_Value = m_Max
    If m_Value <> New_Value Then
        m_Value = New_Value
        m_Percent = (m_Value - m_Min) / (m_Max - m_Min) * 100
        PropertyChanged "Value"
        If Not m_ManualRefresh Then Refresh
        If Format(m_Percent, "0." & Zeros(m_Decimals)) <> Format(m_OldPercent, "0." & Zeros(m_Decimals)) Then
            m_OldPercent = m_Percent
            If Not m_ManualRefresh And Not NoPercent Then Draw
            RaiseEvent Change
        End If
        If m_Value = m_Max Or m_Value = m_Min Then
            If Not m_ManualRefresh Then Draw
            RaiseEvent Change
        End If
    Else
        PropertyChanged "Value"
    End If
End Property
Private Function Zeros(ByVal Count As Byte) As String
    Zeros = ""
    Do Until Count = 0
        Zeros = Zeros & "0"
        Count = Count - 1
    Loop
End Function
'====================================================================================
'====================================================================================
'====================================================================================
'====================================================================================
'====================================================================================
Private Sub BG_Click()
    RaiseEvent Click
End Sub
Private Sub BG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, ScaleX(X - SG.Left, BG.ScaleMode, m_ScaleMode), ScaleY(Y - SG.Top, BG.ScaleMode, m_ScaleMode))
End Sub
Private Sub BG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, ScaleX(X - SG.Left, BG.ScaleMode, m_ScaleMode), ScaleY(Y - SG.Top, BG.ScaleMode, m_ScaleMode))
End Sub
Private Sub BG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, ScaleX(X - SG.Left, BG.ScaleMode, m_ScaleMode), ScaleY(Y - SG.Top, BG.ScaleMode, m_ScaleMode))
End Sub
Private Sub SG_Click()
    RaiseEvent Click
End Sub
Private Sub SG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, ScaleX(X, BG.ScaleMode, m_ScaleMode), ScaleY(Y, BG.ScaleMode, m_ScaleMode))
End Sub
Private Sub SG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, ScaleX(X, BG.ScaleMode, m_ScaleMode), ScaleY(Y, BG.ScaleMode, m_ScaleMode))
End Sub
Private Sub SG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, ScaleX(X, BG.ScaleMode, m_ScaleMode), ScaleY(Y, BG.ScaleMode, m_ScaleMode))
End Sub
Private Sub SG_Resize()
    m_ScaleHeight = SG.ScaleHeight
    m_ScaleWidth = SG.ScaleWidth
End Sub
Private Sub UserControl_InitProperties()
    m_BackAutoColor = m_def_BackAutoColor
    m_BackColor = m_def_BackColor
    m_BackColor2 = m_def_BackColor2
    m_BackColor3 = m_def_BackColor3
    m_BackColor4 = m_def_BackColor4
    m_BackColors = m_def_BackColors
    m_BarAutoColor = m_def_BackAutoColor
    m_BarColor = m_def_BarColor
    m_BarColor2 = m_def_BarColor2
    m_BarColor3 = m_def_BarColor3
    m_BarColor4 = m_def_BarColor4
    m_BarColors = m_def_BarColors
    m_BarStyle = m_def_BarStyle
    m_BarStyleFullOnly = m_def_BarStyleFullOnly
    m_BorderColor = m_def_BorderColor
    m_BorderStyle = m_def_BorderStyle
    m_BorderWidth = m_def_BorderWidth
    m_Custom = m_def_Custom
    m_CustomText = m_def_CustomText
    m_Decimals = m_def_Decimals
    m_Direction = m_def_Direction
    Set m_Font = Ambient.Font
    Set SG.Font = Ambient.Font
    Set FG.Font = Ambient.Font
    m_FormColor = m_def_FormColor
    m_ManualRefresh = m_def_ManualRefresh
    m_Max = m_def_Max
    m_Min = m_def_Min
    m_NoPercent = m_def_NoPercent
    m_Percent = m_def_Percent
    m_PercentAfter = m_def_PercentAfter
    m_PercentAlign = m_def_PercentAlign
    m_PercentBefore = m_def_PercentBefore
    m_PercentColorBar = m_def_PercentColorBar
    m_PercentColorShadow = m_def_PercentColorShadow
    m_ScaleMode = m_def_ScaleMode
    m_ShadowColor = m_def_ShadowColor
    m_ShadowColor2 = m_def_ShadowColor2
    m_ShadowColor3 = m_def_ShadowColor3
    m_ShadowColor4 = m_def_ShadowColor4
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackAutoColor = PropBag.ReadProperty("BackAutoColor", m_def_BackAutoColor)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_BackColor2 = PropBag.ReadProperty("BackColor2", m_def_BackColor2)
    m_BackColor3 = PropBag.ReadProperty("BackColor3", m_def_BackColor3)
    m_BackColor4 = PropBag.ReadProperty("BackColor4", m_def_BackColor4)
    m_BackColors = PropBag.ReadProperty("BackColors", m_def_BackColors)
    m_BarAutoColor = PropBag.ReadProperty("BackAutoColor", m_def_BackAutoColor)
    m_BarColor = PropBag.ReadProperty("BarColor", m_def_BarColor)
    m_BarColor2 = PropBag.ReadProperty("BarColor2", m_def_BarColor2)
    m_BarColor3 = PropBag.ReadProperty("BarColor3", m_def_BarColor3)
    m_BarColor4 = PropBag.ReadProperty("BarColor4", m_def_BarColor4)
    m_BarColors = PropBag.ReadProperty("BarColors", m_def_BarColors)
    m_BarStyle = PropBag.ReadProperty("BarStyle", m_def_BarStyle)
    m_BarStyleFullOnly = PropBag.ReadProperty("BarStyleFullOnly", m_def_BarStyleFullOnly)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_BorderWidth = PropBag.ReadProperty("BorderWidth", m_def_BorderWidth)
    m_Custom = PropBag.ReadProperty("Custom", m_def_Custom)
    m_CustomText = PropBag.ReadProperty("CustomText", m_def_CustomText)
    m_Decimals = PropBag.ReadProperty("Decimals", m_def_Decimals)
    m_Direction = PropBag.ReadProperty("Direction", m_def_Direction)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_FormColor = PropBag.ReadProperty("FormColor", m_def_FormColor)
    m_ManualRefresh = PropBag.ReadProperty("ManualRefresh", m_def_ManualRefresh)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_NoPercent = PropBag.ReadProperty("NoPercent", m_def_NoPercent)
    m_Percent = PropBag.ReadProperty("Percent", m_def_Percent)
    m_PercentAfter = PropBag.ReadProperty("PercentAfter", m_def_PercentAfter)
    m_PercentAlign = PropBag.ReadProperty("PercentAlign", m_def_PercentAlign)
    m_PercentBefore = PropBag.ReadProperty("PercentBefore", m_def_PercentBefore)
    m_PercentColorBar = PropBag.ReadProperty("PercentColorBar", m_def_PercentColorBar)
    m_PercentColorShadow = PropBag.ReadProperty("PercentColorShadow", m_def_PercentColorShadow)
    m_ScaleMode = PropBag.ReadProperty("ScaleMode", m_def_ScaleMode)
    m_ShadowColor = PropBag.ReadProperty("ShadowColor", m_def_ShadowColor)
    m_ShadowColor2 = PropBag.ReadProperty("ShadowColor2", m_def_ShadowColor2)
    m_ShadowColor3 = PropBag.ReadProperty("ShadowColor3", m_def_ShadowColor3)
    m_ShadowColor4 = PropBag.ReadProperty("ShadowColor4", m_def_ShadowColor4)
    m_Value = PropBag.ReadProperty("Value", 0)
    Set SG.Font = m_Font
    Set FG.Font = m_Font
    Select Case m_Direction
        Case 0, 3
            FG.ForeColor = m_PercentColorBar
            SG.ForeColor = m_PercentColorShadow
        Case 1, 2
            SG.ForeColor = m_PercentColorBar
            FG.ForeColor = m_PercentColorShadow
    End Select
    Refresh
    Draw
End Sub
Private Sub UserControl_Resize()
    Dim InWidth As Integer, InHeight As Integer, ScWidth As Integer, ScHeight As Integer
    On Error Resume Next
    'background to entire control
    BG.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    'get right ScaleWidth and ScaleHeight
    ScWidth = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, BG.ScaleMode)
    ScHeight = UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, BG.ScaleMode)
    'size for the progress bar itself
    InWidth = ScWidth - m_BorderWidth * 2
    InHeight = ScHeight - m_BorderWidth * 2
    'check the counted size: if too small, size to entire control
    If InWidth < 1 Then InWidth = ScWidth
    If InHeight < 1 Then InHeight = ScHeight
    'shadowed bar to correct position
    SG.Move (ScWidth - InWidth) / 2, (ScHeight - InHeight) / 2, InWidth, InHeight
    'check direction and size bar accordingly
    Select Case m_Direction
        Case 0 'moving left
            FG.Move -1, 0, (m_Value - m_Min) / (m_Max - m_Min) * (SG.ScaleWidth + 1), SG.ScaleHeight
        Case 1 'moving up
            FG.Move 0, -1, SG.ScaleWidth, ((m_Max - m_Min) - (m_Value - m_Min)) / (m_Max - m_Min) * (SG.ScaleHeight + 1)
        Case 2 'moving right
            FG.Move -1, 0, ((m_Max - m_Min) - (m_Value - m_Min)) / (m_Max - m_Min) * (SG.ScaleWidth + 1), SG.ScaleHeight
        Case 3 'moving down
            FG.Move 0, -1, SG.ScaleWidth, (m_Value - m_Min) / (m_Max - m_Min) * (SG.ScaleHeight + 1)
    End Select
    Bar.Move 0, 0, SG.Width, SG.Height
    Shadow.Move 0, 0, SG.Width, SG.Height
    DrawCorner
    DrawBackground
    DrawBarAndShadow
    If Not m_ManualRefresh Then Draw
    RaiseEvent Resize
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackAutoColor", m_BackAutoColor, m_def_BackAutoColor)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("BackColor2", m_BackColor2, m_def_BackColor2)
    Call PropBag.WriteProperty("BackColor3", m_BackColor3, m_def_BackColor3)
    Call PropBag.WriteProperty("BackColor4", m_BackColor4, m_def_BackColor4)
    Call PropBag.WriteProperty("BackColors", m_BackColors, m_def_BackColors)
    Call PropBag.WriteProperty("BarAutoColor", m_BarAutoColor, m_def_BarAutoColor)
    Call PropBag.WriteProperty("BarColor", m_BarColor, m_def_BarColor)
    Call PropBag.WriteProperty("BarColor2", m_BarColor2, m_def_BarColor2)
    Call PropBag.WriteProperty("BarColor3", m_BarColor3, m_def_BarColor3)
    Call PropBag.WriteProperty("BarColor4", m_BarColor4, m_def_BarColor4)
    Call PropBag.WriteProperty("BarColors", m_BarColors, m_def_BarColors)
    Call PropBag.WriteProperty("BarStyle", m_BarStyle, m_def_BarStyle)
    Call PropBag.WriteProperty("BarStyleFullOnly", m_BarStyleFullOnly, m_def_BarStyleFullOnly)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("BorderWidth", m_BorderWidth, m_def_BorderWidth)
    Call PropBag.WriteProperty("Custom", m_Custom, m_def_Custom)
    Call PropBag.WriteProperty("CustomText", m_CustomText, m_def_CustomText)
    Call PropBag.WriteProperty("Decimals", m_Decimals, m_def_Decimals)
    Call PropBag.WriteProperty("Direction", m_Direction, m_def_Direction)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("FormColor", m_FormColor, m_def_FormColor)
    Call PropBag.WriteProperty("ManualRefresh", m_ManualRefresh, m_def_ManualRefresh)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("NoPercent", m_NoPercent, m_def_NoPercent)
    Call PropBag.WriteProperty("Percent", m_Percent, m_def_Percent)
    Call PropBag.WriteProperty("PercentAfter", m_PercentAfter, m_def_PercentAfter)
    Call PropBag.WriteProperty("PercentAlign", m_PercentAlign, m_def_PercentAlign)
    Call PropBag.WriteProperty("PercentBefore", m_PercentBefore, m_def_PercentBefore)
    Call PropBag.WriteProperty("PercentColorBar", m_PercentColorBar, m_def_PercentColorBar)
    Call PropBag.WriteProperty("PercentColorShadow", m_PercentColorShadow, m_def_PercentColorShadow)
    Call PropBag.WriteProperty("ScaleMode", m_ScaleMode, m_def_ScaleMode)
    Call PropBag.WriteProperty("ShadowColor", m_ShadowColor, m_def_ShadowColor)
    Call PropBag.WriteProperty("ShadowColor2", m_ShadowColor2, m_def_ShadowColor2)
    Call PropBag.WriteProperty("ShadowColor3", m_ShadowColor3, m_def_ShadowColor3)
    Call PropBag.WriteProperty("ShadowColor4", m_ShadowColor4, m_def_ShadowColor4)
    Call PropBag.WriteProperty("Value", m_Value, 0)
End Sub
