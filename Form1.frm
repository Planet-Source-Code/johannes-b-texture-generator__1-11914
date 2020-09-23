VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "COOL!"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   7440
      Top             =   3600
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4335
      Left            =   -360
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   421
      TabIndex        =   3
      Top             =   1320
      Width           =   6375
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton Command1 
         Caption         =   "DRAW"
         Height          =   375
         Left            =   4680
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Auto redraw"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   120
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Auto refresh"
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   120
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Text            =   "5"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "3"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "2"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "1"
         Top             =   840
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "API calls"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Color delay"
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   0
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label3 
         Caption         =   "B"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "G"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "R"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   135
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim XX, YY, F1, F2, F3, FF1, FF2, FF3 As Integer
Dim DL As Integer

Private Sub Check3_Click()
If Check3.Value = 1 Then
Picture1.AutoRedraw = True
Check2.Enabled = True
Else
Picture1.AutoRedraw = False
Check2.Value = 0
Check2.Enabled = False
End If
End Sub

Private Sub Command1_Click()
XX = -1
YY = -1
Timer1.Interval = "100"
Form1.Caption = "DRAWING..."


End Sub

Private Sub Command2_Click()
Form1.Caption = "COOL!"
Timer1.Interval = "0"
End Sub

Private Sub Form_Resize()
Picture1.Width = Form1.ScaleWidth
Picture1.Height = Form1.ScaleHeight - 90
End Sub

Private Sub Timer1_Timer()
On Error Resume Next


For Counter = 1 To 5000

If DL > Text4.Text Then
'COLOR CYCLING
If Option1.Value = True Then
If F1 >= 270 Then F11 = 1
If F2 >= 270 Then F22 = 1
If F3 >= 270 Then F33 = 1

If F1 <= 20 Then F11 = 0
If F2 <= 20 Then F22 = 0
If F3 <= 20 Then F33 = 0
End If


If F11 = 0 Then
F1 = F1 + 3 * Text1.Text
Else
F1 = F1 - 5 * Text1.Text
End If



If F22 = 0 Then
F2 = F2 + 4 * Text2.Text
Else
F2 = F2 - 4 * Text2.Text
End If



If F33 = 0 Then
F3 = F3 + 5 * Text3.Text
Else
F3 = F3 - 3 * Text3.Text
End If

DL = 0






End If

DL = DL + 1

'DRAW



If Check1.Value = 1 Then
'API
SetPixelV Picture1.hDC, XX, YY, RGB(F1, F2, F3)
Else
'VB STANDARD
Picture1.ForeColor = RGB(F1, F2, F3)
Picture1.PSet (XX, YY)
End If


XX = XX + 1

If XX > Picture1.ScaleWidth Then
XX = -1
YY = YY + 1
End If


If YY > Picture1.ScaleHeight Then
Form1.Caption = "COOL!"
Timer1.Interval = "0"
End If


Next


If Check2.Value = 1 Then
Picture1.Refresh
End If


End Sub



