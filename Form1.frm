VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox kep 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0C0C0&
      Height          =   1665
      Index           =   0
      Left            =   480
      ScaleHeight     =   28.84
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   76.994
      TabIndex        =   0
      Top             =   930
      Width           =   4395
      Begin VB.PictureBox sk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   0
         Left            =   3060
         MousePointer    =   8  'Size NW SE
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   1.852
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   1.852
         TabIndex        =   9
         Top             =   420
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.PictureBox sk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   7
         Left            =   3570
         MousePointer    =   8  'Size NW SE
         Picture         =   "Form1.frx":00EA
         ScaleHeight     =   1.852
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   1.852
         TabIndex        =   8
         Top             =   900
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.PictureBox sk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   6
         Left            =   3300
         MousePointer    =   7  'Size N S
         Picture         =   "Form1.frx":01D4
         ScaleHeight     =   1.852
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   1.852
         TabIndex        =   7
         Top             =   900
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.PictureBox sk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   5
         Left            =   3060
         MousePointer    =   6  'Size NE SW
         Picture         =   "Form1.frx":02BE
         ScaleHeight     =   1.852
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   1.852
         TabIndex        =   6
         Top             =   870
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.PictureBox sk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   4
         Left            =   3540
         MousePointer    =   9  'Size W E
         Picture         =   "Form1.frx":03A8
         ScaleHeight     =   1.852
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   1.852
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.PictureBox sk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   3
         Left            =   3060
         MousePointer    =   9  'Size W E
         Picture         =   "Form1.frx":0492
         ScaleHeight     =   1.852
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   1.852
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.PictureBox sk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   2
         Left            =   3540
         MousePointer    =   6  'Size NE SW
         Picture         =   "Form1.frx":057C
         ScaleHeight     =   1.852
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   1.852
         TabIndex        =   3
         Top             =   420
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.PictureBox sk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   1
         Left            =   3270
         MousePointer    =   7  'Size N S
         Picture         =   "Form1.frx":0666
         ScaleHeight     =   1.852
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   1.852
         TabIndex        =   2
         Top             =   390
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   525
         Index           =   0
         Left            =   840
         ScaleHeight     =   8.731
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   14.023
         TabIndex        =   1
         Top             =   330
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Shape nk 
         BorderColor     =   &H00404040&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   900
         Top             =   1140
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Shape Shape 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   0
         Left            =   2280
         Top             =   810
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Shape frame 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   1830
         Top             =   30
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.Menu newobj 
      Caption         =   "New object"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim movingPic As Boolean
Dim aX, aY, topx, topy
Dim Actobject()
Dim Tempobject, picindex As Integer
Sub NewObject()
    k = Pic.Count
    Load Pic(k)
    Pic(k).Left = 5
    Pic(k).Top = 5

    Pic(k).Left = 10
    Pic(k).Width = 20
    Pic(k).Top = 10
    Pic(k).Height = 20
    Pic(k).BackColor = &H808000
    
    Pic(k).BorderStyle = 1
    Pic(k).Visible = True
    Pic_Paint k + 0
    'Down Pic(k), 0, 0, 0
    'Up Pic(k), 0, 0, 0
End Sub
Sub torsk()
' remove shapes
Do While Shape.Count > 1
    Unload Shape(Shape.Count - 1)
    Unload frame(frame.Count - 1)
Loop
Do While sk.Count > 8
    Unload sk(sk.Count - 1)
Loop

End Sub

Sub Up(obj, X As Single, Y As Single, Shift%)
        van = -1
        If Actobject(0) Is Nothing = False Then
            For i = 0 To UBound(Actobject)
                If Actobject(i).Name = obj.Name And Actobject(i).Index = obj.Index Then
                    van = i
                End If
            Next
        End If
    '''''''''''''''''''''''
    If Actobject(0) Is Nothing Then
        Set Actobject(0) = obj
            torsk
    Else
        If Shift = 0 Then
            If van = -1 Then
                ReDim Actobject(0)
                Set Actobject(0) = obj
                torsk
            End If
        Else
            If van = -1 Then
                ReDim Preserve Actobject(UBound(Actobject) + 1)
                Set Actobject(UBound(Actobject)) = obj
                ujsk
                Shape(UBound(Actobject)).Move Actobject(UBound(Actobject)).Left, Actobject(UBound(Actobject)).Top, Actobject(UBound(Actobject)).Width, Actobject(UBound(Actobject)).Height
            Else
                If UBound(Actobject) > 0 Then visszavesz van + 0
            End If
        End If
    End If
    For i = 0 To UBound(Actobject)
            If Shape(i).Top < 0 Then Shape(i).Top = 2
            If Shape(i).Left < 0 Then Shape(i).Left = 2
            If Shape(i).Width < 10 Then Shape(i).Width = 10
            If Shape(i).Height < 10 Then Shape(i).Height = 10
'            Stop
        e = Round(2 * Shape(i).Left / 2): If e / 2 <> e \ 2 Then e = e + 1
        f = Round(2 * Shape(i).Top / 2): If f / 2 <> f \ 2 Then f = f + 1
            Shape(i).Left = e
            Shape(i).Top = f
            
            
            Actobject(i).Move Shape(i).Left, Shape(i).Top, Shape(i).Width, Shape(i).Height
nem:
            Actobject(i).Visible = True
            
            FocusPic Actobject(i), i + 0
            On Error Resume Next
            Actobject(i).SetFocus
            Actobject(i).ZOrder 0
            
    Next
On Error GoTo 0
Err.Clear
End Sub
Sub FocusPic(mit, ii As Integer)
iii = ii * 7 + ii
frame(ii).Move mit.Left - 1, mit.Top - 1, mit.Width + 2, mit.Height + 2
sk(iii).Left = mit.Left - sk(0).Width: sk(iii).Top = mit.Top - sk(0).Height
sk(iii + 1).Left = mit.Left + (mit.Width - sk(0).Width) / 2: sk(iii + 1).Top = mit.Top - sk(1).Height
sk(iii + 2).Left = mit.Left + mit.Width: sk(iii + 2).Top = mit.Top - sk(0).Height

sk(iii + 3).Left = mit.Left - sk(0).Width: sk(iii + 3).Top = mit.Top + (mit.Height - sk(0).Height) / 2
sk(iii + 4).Left = mit.Left + mit.Width: sk(iii + 4).Top = mit.Top + (mit.Height - sk(0).Height) / 2

sk(iii + 5).Left = mit.Left - sk(0).Width: sk(iii + 5).Top = mit.Top + mit.Height
sk(iii + 6).Left = mit.Left + (mit.Width - sk(0).Width) / 2: sk(iii + 6).Top = mit.Top + mit.Height
sk(iii + 7).Left = mit.Left + mit.Width: sk(iii + 7).Top = mit.Top + mit.Height
'mit.Cls


PicPrint mit.Index

frame(ii).Visible = True
For i = ii * 7 + ii To ii * 7 + ii + 7
    If mit.Tag = 0 Then
        If i = ii * 7 + ii + 0 Then GoTo nemkell
        If i = ii * 7 + ii + 1 Then GoTo nemkell
        If i = ii * 7 + ii + 2 Then GoTo nemkell
        If i = ii * 7 + ii + 5 Then GoTo nemkell
        If i = ii * 7 + ii + 6 Then GoTo nemkell
        If i = ii * 7 + ii + 7 Then GoTo nemkell
    End If
    If mit.Tag = 1 Then
        If i = ii * 7 + ii + 0 Then GoTo nemkell
        If i = ii * 7 + ii + 2 Then GoTo nemkell
        If i = ii * 7 + ii + 3 Then GoTo nemkell
        If i = ii * 7 + ii + 4 Then GoTo nemkell
        If i = ii * 7 + ii + 5 Then GoTo nemkell
        If i = ii * 7 + ii + 7 Then GoTo nemkell
    End If
    
    sk(i).Visible = True
    sk(i).ZOrder 0
nemkell:
Next

End Sub

Private Sub Form_Load()
ReDim Actobject(0)
Set Actobject(0) = Nothing

End Sub
Sub ujsk()
On Error GoTo errlabel
k = UBound(Actobject)
'GoTo ki
Load Shape(k)
Load frame(k)
i = k * 7 + k
Load sk(i): Set sk(i) = sk(0): i = i + 1
Load sk(i): Set sk(i) = sk(1): i = i + 1
Load sk(i): Set sk(i) = sk(2): i = i + 1
Load sk(i): Set sk(i) = sk(3): i = i + 1
Load sk(i): Set sk(i) = sk(4): i = i + 1
Load sk(i): Set sk(i) = sk(5): i = i + 1
Load sk(i): Set sk(i) = sk(6): i = i + 1
Load sk(i): Set sk(i) = sk(7)
exi:
On Error GoTo 0
Exit Sub
errlabel:
Err.Clear
Resume exi
End Sub

Private Sub Form_Resize()
'TabControl1.Move 90, 30, Me.Width - 120, Me.Height - 300
kep(0).Move 60, 120, Me.Width - 240, Me.Height - kep(0).Top - 120
kep(0).Cls
gridel 0

End Sub
Sub gridel(k)
'painting grid points
For i = 0 To kep(0).Width / 56.7 Step 2
    For j = 0 To kep(0).Height / 56.7 Step 2
        kep(0).PSet (i, j)
    Next
Next
End Sub

Private Sub kep_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'paste new object
If Shift = 2 And KeyCode = 86 Then
    If Tempobject > 0 Then

    NewObject
    End If
End If

End Sub

Private Sub kep_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To UBound(Actobject)
    frame(i).Visible = False
    Shape(i).Visible = 0
Next
    ReDim Actobject(0)
    Set Actobject(0) = Nothing
    torsk
    For i = 0 To sk.Count - 1
        sk(i).Visible = False
    Next
    topx = X: topy = Y
    nk.Move X, Y, 0, 0: nk.Visible = True


End Sub

Private Sub kep_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    If X > topx And Y > topy Then
        nk.Width = X - topx
        nk.Height = Y - topy
    ElseIf X > topx And Y < topy Then
        nk.Width = X - topx
        nk.Top = Y
        nk.Height = topy - Y
    ElseIf X < topx And Y > topy Then
        nk.Left = X
        nk.Width = topx - X
        nk.Height = Y - topy
    Else
        nk.Left = X
        nk.Width = topx - X
        nk.Top = Y
        nk.Height = topy - Y
    End If
End If
On Error GoTo 0
Err.Clear

End Sub

Private Sub kep_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
nk.Visible = False
LoadObj nk.Left, nk.Top, nk.Left + nk.Width, nk.Height + nk.Top

End Sub
Sub LoadObj(x0, y0, x1, y1)
ReDim Actobject(0)
Set Actobject(0) = Nothing
'For i = 1 To pic.Count - 1
    For Each kk In Pic
    j = kk.Index
    If j = 0 Then GoTo kovi1
    If Pic(j).MousePointer = 11 Then GoTo kovi1
    If (Pic(j).Left >= x0 And Pic(j).Left <= x1) And (Pic(j).Top >= y0 And Pic(j).Top <= y1) Then
        AddObj Pic(j)
        GoTo kovi1
    End If
    If (Pic(j).Left + Pic(j).Width >= x0 And Pic(j).Left + Pic(j).Width <= x1) And (Pic(j).Top + Pic(j).Height >= y0 And Pic(j).Top + Pic(j).Height <= y1) Then
        AddObj Pic(j)
        GoTo kovi1
    End If
    If (Pic(j).Left + Pic(j).Width >= x0 And Pic(j).Left + Pic(j).Width <= x1) And (Pic(j).Top >= y0 And Pic(j).Top <= y1) Then
        AddObj Pic(j)
        GoTo kovi1
    End If
    If (Pic(j).Left >= x0 And Pic(j).Left <= x1) And (Pic(j).Top + Pic(j).Height >= y0 And Pic(j).Top + Pic(j).Height <= y1) Then
        AddObj Pic(j)
        GoTo kovi1
    End If
kovi1:
Next

End Sub
Sub AddObj(obj)
If Actobject(0) Is Nothing Then
            Set Actobject(0) = obj
Else
            ReDim Preserve Actobject(UBound(Actobject) + 1)
            Set Actobject(UBound(Actobject)) = obj
            ujsk
End If
FocusPic Actobject(UBound(Actobject)), UBound(Actobject)
End Sub
Private Sub newobj_Click()
NewObject
End Sub

Private Sub Pic_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 67 Then
    Tempobject = Index
End If

End Sub

Private Sub Pic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Down Pic(Index), X, Y, Shift
End If

End Sub

Private Sub Pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Shift = 0 Then Moving X, Y

End Sub

Private Sub Pic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Up Pic(Index), X, Y, Shift

End Sub

Sub Moving(jx, jy)
If movingPic Then Exit Sub
movingPic = True
Me.Enabled = False
On Error Resume Next
dx = jx - aX
dy = jy - aY
If dx = 0 And dy = 0 Then GoTo kesz
obin = 0
If picindex = -1 Then
    For obin = 0 To UBound(Actobject)
        Shape(obin).Move Actobject(obin).Left + dx, Actobject(obin).Top + dy
        DoEvents
    Next
    GoTo kesz
End If
egy:
'If picindex = -1 Then
'    Shape(obin).Move Actobject(obin).Left + dx, Actobject(obin).Top + dy
If picindex = 0 Then
    Shape(obin).Move Actobject(obin).Left + dx, Actobject(obin).Top + dy, Actobject(obin).Width - dx, Actobject(obin).Height - dy
ElseIf picindex = 1 Then 'fel
    Shape(obin).Move Actobject(obin).Left, Actobject(obin).Top + dy, Actobject(obin).Width, Actobject(obin).Height - dy
ElseIf picindex = 2 Then
    Shape(obin).Move Actobject(obin).Left, Actobject(obin).Top + dy, Actobject(obin).Width + dx, Actobject(obin).Height - dy
ElseIf picindex = 3 Then
    Shape(obin).Move Actobject(obin).Left + dx, Actobject(obin).Top, Actobject(obin).Width - dx, Actobject(obin).Height
ElseIf picindex = 4 Then
    Shape(obin).Move Actobject(obin).Left, Actobject(obin).Top, Actobject(obin).Width + dx, Actobject(obin).Height
ElseIf picindex = 5 Then
    Shape(obin).Move Actobject(obin).Left + dx, Actobject(obin).Top, Actobject(obin).Width - dx, Actobject(obin).Height + dy
ElseIf picindex = 6 Then
    Shape(obin).Move Actobject(obin).Left, Actobject(obin).Top, Actobject(obin).Width, Actobject(obin).Height + dy
Else
    Shape(obin).Move Actobject(obin).Left, Actobject(obin).Top, Actobject(obin).Width + dx, Actobject(obin).Height + dy
End If
kesz:
For i = 0 To sk.Count
    sk(i).Visible = False
Next
On Error GoTo 0
Err.Clear
Me.Enabled = True
movingPic = False
End Sub

Private Sub Pic_Paint(Index As Integer)
    PicPrint Index

End Sub
Sub PicPrint(idd%)
    tem$ = "Obj"
    ex = (Pic(idd).Width - Pic(idd).TextWidth(tem$)) / 2
    fx = (Pic(idd).Height - Pic(idd).TextHeight(tem$)) / 2
    Pic(idd).Cls
    Pic(idd).CurrentX = ex
    Pic(idd).CurrentY = fx
    Pic(idd).Print tem$
    
End Sub

Sub DelPic()
For i = 0 To Shape.Count - 1
    Shape(i).Visible = False
    frame(i).Visible = False
Next
For i = 0 To sk.Count - 1
    sk(i).Visible = False
Next

End Sub
Sub Down(obj, X As Single, Y As Single, Shift As Integer)
frame(0).Visible = False
aX = X: aY = Y
picindex = -1
DelPic
        van = -1
        If Actobject(0) Is Nothing = False Then
            For i = 0 To UBound(Actobject)
                If Actobject(i).Name = obj.Name And Actobject(i).Index = obj.Index Then
                    van = i
                End If
            Next
        End If
If Shift = 0 And (UBound(Actobject) = 0 Or van = -1) Then
    ReDim Actobject(0)
    Set Actobject(0) = obj
    Actobject(0).Visible = False
    Shape(0).Move Actobject(0).Left, Actobject(0).Top, Actobject(0).Width, Actobject(0).Height
    Shape(0).Visible = True
    Shape(0).ZOrder 0
Else
    If Actobject(0) Is Nothing Then GoTo att
    For i = 0 To UBound(Actobject)
    Actobject(i).Visible = False
    Shape(i).Move Actobject(i).Left, Actobject(i).Top, Actobject(i).Width, Actobject(i).Height
    Shape(i).Visible = True
    Shape(i).ZOrder 0
    Next
att:
End If




End Sub

Sub visszavesz(mi%)

Actobject(mi%).Visible = True
If mi < UBound(Actobject) Then
    For i = mi To UBound(Actobject) - 1
        Set Actobject(i) = Actobject(i + 1)
         Shape(i).Left = Shape(i + 1).Left
         Shape(i).Top = Shape(i + 1).Top
         Shape(i).Width = Shape(i + 1).Width
         Shape(i).Height = Shape(i + 1).Height
         frame(i).Left = frame(i + 1).Left
         frame(i).Top = frame(i + 1).Top
         frame(i).Width = frame(i + 1).Width
         frame(i).Height = frame(i + 1).Height
    
    Next
End If
''''''
ReDim Preserve Actobject(UBound(Actobject) - 1)
Unload Shape(Shape.Count - 1)
Unload keret(keret.Count - 1)
    Do While sk.Count > 8 * (UBound(Actobject) + 1)
        Unload sk(sk.Count - 1)
    Loop
End Sub

Private Sub sk_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If UBound(Actobject) > 0 Then Exit Sub

aX = X
aY = Y
On Error Resume Next
    
    Shape(0).Move Actobject(0).Left, Actobject(0).Top, Actobject(0).Width, Actobject(0).Height
    Actobject(0).Visible = False

    frame(0).Visible = False
    Shape(0).Visible = True
    Shape(0).ZOrder 0
Err.Clear
On Error GoTo 0
End Sub

Private Sub sk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If UBound(Actobject) > 0 Then Exit Sub
If Button = 1 Then
    picindex = Index
    Moving X, Y
End If
End Sub

Private Sub sk_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If UBound(Actobject) > 0 Then Exit Sub

            Shape(0).Visible = False
            If Shape(i).Top < 0 Then Shape(i).Top = 2
            If Shape(i).Left < 0 Then Shape(i).Left = 2
            If Shape(i).Width < 10 Then Shape(i).Width = 10
            If Shape(i).Height < 10 Then Shape(i).Height = 10

    
    
    Actobject(0).Move Shape(0).Left, Shape(0).Top, Shape(0).Width, Shape(0).Height
    If Val(Actobject(0).Tag) = 5 Then Actobject(0).Width = Actobject(0).Height + 0.39
    
    Actobject(0).Visible = True
    FocusPic Actobject(0), 0
End Sub
