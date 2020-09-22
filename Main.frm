VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bezier Curve Generator"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   657
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox RealtimeCheckbox 
      Caption         =   "Realtime"
      Height          =   255
      Left            =   6600
      TabIndex        =   14
      Top             =   7800
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.HScrollBar DetailScroll 
      Height          =   255
      LargeChange     =   10
      Left            =   6600
      Max             =   1
      Min             =   100
      TabIndex        =   13
      Top             =   7440
      Value           =   1
      Width           =   3135
   End
   Begin VB.TextBox YPosText 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3840
      TabIndex        =   11
      Text            =   "0"
      Top             =   7920
      Width           =   1335
   End
   Begin VB.TextBox XPosText 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3840
      TabIndex        =   10
      Text            =   "0"
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton NextButton 
      Caption         =   ">"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton PreviousButton 
      Caption         =   "<"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton CreateButton 
      Caption         =   "Recreate"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox CreateNumberText 
      Height          =   285
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "2"
      Top             =   7440
      Width           =   735
   End
   Begin VB.PictureBox DisplayPicture 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   120
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   120
      Width           =   9600
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Detail Level:"
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   7440
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   360
      Y1              =   544
      Y2              =   496
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Y Coordinate:"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "X Coordinate:"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   160
      X2              =   160
      Y1              =   544
      Y2              =   496
   End
   Begin VB.Label PointNumberLabel 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   7440
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Point Number:"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Control Points:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7440
      Width           =   1215
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ControlPoints() As BezierPoint
Dim CurrentPoint As Long

Private Sub UpdateDisplay(Optional DrawCurves As Boolean = True)
    Dim Pos As Long, Time As Double, Detail As Double
    Dim OldPoint As BezierPoint, NewPoint As BezierPoint
    PointNumberLabel.Caption = CStr(CurrentPoint + 1)
    XPosText.Text = CStr(ControlPoints(CurrentPoint).X)
    YPosText.Text = CStr(ControlPoints(CurrentPoint).Y)
    
    DisplayPicture.Cls
    DisplayPicture.DrawWidth = 1
    If DrawCurves = True Then
        OldPoint = CalculateBezier(0, ControlPoints)
        Detail = DetailScroll.Value / 1000
        For Time = Detail To 1 Step Detail
            NewPoint = CalculateBezier(Time, ControlPoints)
            DisplayPicture.Line (OldPoint.X, OldPoint.Y)-(NewPoint.X, NewPoint.Y), RGB(150, 150, 150)
            OldPoint = NewPoint
        Next
        NewPoint = CalculateBezier(1, ControlPoints)
        DisplayPicture.Line (OldPoint.X, OldPoint.Y)-(NewPoint.X, NewPoint.Y), RGB(150, 150, 150)
    End If
    
    DisplayPicture.DrawWidth = 5
    For Pos = 0 To UBound(ControlPoints)
        If Pos = CurrentPoint Then
            DisplayPicture.PSet (ControlPoints(Pos).X, ControlPoints(Pos).Y), RGB(255, 0, 0)
        Else
            DisplayPicture.PSet (ControlPoints(Pos).X, ControlPoints(Pos).Y), RGB(150, 150, 150)
        End If
    Next
End Sub

Private Sub CreateButton_Click()
    CurrentPoint = 0
    ReDim ControlPoints(Val(CreateNumberText.Text) - 1)
    UpdateDisplay
End Sub

Private Sub CreateNumberText_LostFocus()
    CreateNumberText.Text = Str(Val(CreateNumberText.Text))
    If Val(CreateNumberText.Text) < 1 Then CreateNumberText.Text = "1"
End Sub

Private Sub DetailScroll_Change()
    UpdateDisplay
End Sub

Private Sub DisplayPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ControlPoints(CurrentPoint).X = CLng(X)
    ControlPoints(CurrentPoint).Y = CLng(Y)
    If RealtimeCheckbox.Value = 1 Then
        UpdateDisplay True
    Else
        UpdateDisplay False
    End If
End Sub

Private Sub DisplayPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button > 0 Then
        ControlPoints(CurrentPoint).X = CLng(X)
        ControlPoints(CurrentPoint).Y = CLng(Y)
        If RealtimeCheckbox.Value = 1 Then
            UpdateDisplay True
        Else
            UpdateDisplay False
        End If
    End If
End Sub

Private Sub DisplayPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ControlPoints(CurrentPoint).X = CLng(X)
    ControlPoints(CurrentPoint).Y = CLng(Y)
    UpdateDisplay True
End Sub

Private Sub Form_Load()
    ReDim ControlPoints(1)
    CurrentPoint = 0
    UpdateDisplay
End Sub

Private Sub NextButton_Click()
    CurrentPoint = CurrentPoint + 1
    If CurrentPoint > UBound(ControlPoints) Then CurrentPoint = 0
    UpdateDisplay
End Sub

Private Sub PreviousButton_Click()
    CurrentPoint = CurrentPoint - 1
    If CurrentPoint < 0 Then CurrentPoint = UBound(ControlPoints)
    UpdateDisplay
End Sub

Private Sub XPosText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        ControlPoints(CurrentPoint).X = Val(XPosText.Text)
        UpdateDisplay
    End If
End Sub

Private Sub XPosText_LostFocus()
    ControlPoints(CurrentPoint).X = Val(XPosText.Text)
    UpdateDisplay
End Sub

Private Sub YPosText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        ControlPoints(CurrentPoint).Y = Val(YPosText.Text)
        UpdateDisplay
    End If
End Sub

Private Sub YPosText_LostFocus()
    ControlPoints(CurrentPoint).Y = Val(YPosText.Text)
    UpdateDisplay
End Sub
