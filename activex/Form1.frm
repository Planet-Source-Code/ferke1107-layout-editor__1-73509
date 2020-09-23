VERSION 5.00
Object = "*\AProject1.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin Project1.EditForm EditForm1 
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   570
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4260
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1710
      TabIndex        =   2
      Text            =   "Caption"
      Top             =   90
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "delete object"
      Height          =   375
      Left            =   3180
      TabIndex        =   1
      Top             =   60
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "add object"
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim akt%()
Private Sub Command1_Click()
'sample color
EditForm1.NewObject Text1, vbBlue
End Sub

Private Sub Command2_Click()
EditForm1.DeleteObject akt
End Sub

Private Sub editform1_SelectedChange(index() As Integer)
'Me.Caption = index
k = -1
ReDim akt(0)
For i = 0 To UBound(index)
    k = k + 1
    ReDim Preserve akt(k)
    akt(k) = index(i)
Next
End Sub

Private Sub Form_Resize()
EditForm1.Width = Me.Width - 300
EditForm1.Height = Me.Height - EditForm1.Top - 800

End Sub
