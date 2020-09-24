VERSION 5.00
Object = "{B521A17A-83E1-4BEF-B7F1-7E95A209E13D}#26.0#0"; "MunPBDeluxe.ocx"
Begin VB.Form MunPBDexample1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minun Progress Bar Deluxe"
   ClientHeight    =   3495
   ClientLeft      =   4665
   ClientTop       =   3675
   ClientWidth     =   4575
   Icon            =   "MunPBDexample1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton FullUp 
      Caption         =   "Go full"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   3735
   End
   Begin MinunProgressBarDeluxe.MinunPBDeluxe PB 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      Custom          =   -1  'True
      CustomText      =   "Finished!"
      Decimals        =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   30000
      PercentAlign    =   5
      PercentBefore   =   "Loading program..."
      ScaleMode       =   3
   End
   Begin MinunProgressBarDeluxe.MinunPBDeluxe PB 
      Height          =   2775
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   4895
      BackColor2      =   4210688
      BackColor3      =   8421376
      BackColor4      =   12632064
      BarColor        =   192
      BarColor2       =   128
      BarColor3       =   192
      BarColor4       =   8421631
      Direction       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   30000
      PercentAlign    =   3
      PercentColorBar =   8438015
      PercentColorShadow=   16512
      ScaleMode       =   3
      ShadowColor     =   12640511
      ShadowColor2    =   8438015
      ShadowColor3    =   12640511
      ShadowColor4    =   12648447
   End
   Begin MinunProgressBarDeluxe.MinunPBDeluxe PB 
      Height          =   735
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1296
      BackColor2      =   12632319
      BackColor3      =   8421631
      BackColor4      =   65535
      BarColor        =   16744576
      BarColor2       =   8388736
      BarColor3       =   16711935
      BarColor4       =   8388736
      BorderColor     =   33023
      BorderWidth     =   8
      Custom          =   -1  'True
      CustomText      =   "Finished!"
      Decimals        =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   3000
      PercentAlign    =   6
      PercentColorBar =   8388608
      PercentColorShadow=   128
      ScaleMode       =   3
      ShadowColor     =   16761087
      ShadowColor2    =   8388736
      ShadowColor3    =   16711935
      ShadowColor4    =   8388736
   End
End
Attribute VB_Name = "MunPBDexample1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim A As Integer
    For A = 0 To 2
        PB(A).BackColors = mpbdFourColors
    Next A
End Sub
Private Sub FullUp_Click()
    PB(1).Value = 0
    PB(2).Value = 0
    Do While PB(1).Value < PB(1).Max
        PB(1).Value = PB(1).Value + 1
        PB(0).Value = (PB(1).Value + PB(2).Value) / (PB(1).Max + PB(2).Max) * PB(0).Max
    Loop
    Do While PB(2).Value < PB(2).Max
        PB(2).Value = PB(2).Value + 1
        PB(0).Value = (PB(1).Value + PB(2).Value) / (PB(1).Max + PB(2).Max) * PB(0).Max
    Loop
End Sub
Private Sub PB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    PB_MouseMove Index, Button, Shift, X, Y
End Sub
Private Sub PB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    PB(Index).BackColors = mpbdThreeColors
    If Button <> 1 Then Exit Sub
    Select Case PB(Index).Direction
        Case 0
            PB(Index).Value = X / PB(Index).ScaleWidth * (PB(Index).Max - PB(Index).Min) + PB(Index).Min
        Case 1
            PB(Index).Value = (PB(Index).ScaleHeight - Y) / PB(Index).ScaleHeight * (PB(Index).Max - PB(Index).Min) + PB(Index).Min
        Case 2
            PB(Index).Value = (PB(Index).ScaleWidth - X) / PB(Index).ScaleWidth * (PB(Index).Max - PB(Index).Min) + PB(Index).Min
        Case 3
            PB(Index).Value = Y / PB(Index).ScaleHeight * (PB(Index).Max - PB(Index).Min) + PB(Index).Min
    End Select
End Sub
Private Sub PB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    PB_MouseMove Index, Button, Shift, X, Y
End Sub
