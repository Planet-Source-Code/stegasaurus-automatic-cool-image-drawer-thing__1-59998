VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Automatic Cool Image Drawer"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command13 
      Caption         =   "-"
      Height          =   255
      Left            =   1200
      TabIndex        =   27
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      Caption         =   "+"
      Height          =   255
      Left            =   840
      TabIndex        =   26
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Make Central"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Save Image"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   23
      Text            =   "5"
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Restart"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Stop"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "-"
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "+"
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "-"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   12
      Text            =   "0"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Text            =   "0"
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   2160
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Drawing Dot"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Central Dot"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      DrawWidth       =   5
      FillStyle       =   0  'Solid
      Height          =   5775
      Left            =   1680
      ScaleHeight     =   5715
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   0
         Top             =   0
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   0
         Shape           =   3  'Circle
         Tag             =   "0,0"
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape Gravity 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Line Width"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Y"
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "X"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Central Dot Position:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Background:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Foreground:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Image Drawer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xspeed As Long
Dim yspeed As Long
Private Sub Check1_Click()
Gravity.Visible = Check1.Value
End Sub
Private Sub Check2_Click()
Shape1.Visible = Check2.Value
End Sub
Private Sub Command1_Click()
cdl.ShowColor
Label3.BackColor = cdl.Color
Picture1.ForeColor = Label3.BackColor
End Sub
Private Sub Command10_Click()
cdl.Filter = "Bitmaps (*.bmp)|*.bmp"
If cdl.FileName = "" Then
    cdl.ShowSave
    If cdl.FileName = "" Then
    Exit Sub
    Else
    SavePicture Picture1.Image, cdl.FileName
    End If
    Else
    intsave = MsgBox("Save over existing file?", _
    vbYesNoCancel + vbExclamation)
    Select Case intsave
    Case vbYes
    SavePicture Picture1.Image, cdl.FileName
    Case vbNo
    cdl.FileName = ""
    cdl.ShowSave
    If cdl.FileName = "" Then
    Exit Sub
    Else
    SavePicture Picture1.Image, cdl.FileName
    End If
    Case vbCancel
    Exit Sub
    End Select
    End If
End Sub
Private Sub Command11_Click()
Gravity.Top = Picture1.ScaleHeight / 2
Text2.Text = Gravity.Top
Gravity.Left = Picture1.ScaleWidth / 2
Text1.Text = Gravity.Left
End Sub
Private Sub Command12_Click()
Text3.Text = Val(Text3.Text) + 1
End Sub
Private Sub Command13_Click()
Text3.Text = Val(Text3.Text) - 1
End Sub
Private Sub Command2_Click()
cdl.ShowColor
Label5.BackColor = cdl.Color
Picture1.BackColor = Label5.BackColor
End Sub
Private Sub Command3_Click()
Text1.Text = Val(Text1.Text) + 50
End Sub
Private Sub Command4_Click()
Text1.Text = Val(Text1.Text) - 50
End Sub
Private Sub Command5_Click()
Text2.Text = Val(Text2.Text) + 50
End Sub
Private Sub Command6_Click()
Text2.Text = Val(Text2.Text) - 50
End Sub
Private Sub Command7_Click()
Timer1.Enabled = True
End Sub
Private Sub Command8_Click()
Timer1.Enabled = False
End Sub
Private Sub Command9_Click()
Timer1.Enabled = False
Picture1.Cls
Shape1.Top = 0
Shape1.Left = 0
xspeed = 0
yspeed = 0
Timer1.Enabled = True
End Sub
Private Sub Form_Load()
Form_Resize
Gravity.Top = Picture1.ScaleHeight / 2
Text2.Text = Gravity.Top
Gravity.Left = Picture1.ScaleWidth / 2
Text1.Text = Gravity.Left
End Sub
Private Sub Form_Resize()
Picture1.Width = Me.ScaleWidth - Picture1.Left
Picture1.Height = Me.ScaleHeight
End Sub
Private Sub Text1_Change()
Gravity.Left = Text1.Text
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Process(KeyAscii)
End Sub
Private Sub Text2_Change()
Gravity.Top = Val(Text2.Text)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = Process(KeyAscii)
End Sub
Private Sub Text3_Change()
Picture1.DrawWidth = Text3.Text
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = Process(KeyAscii)
End Sub
Private Sub Timer1_Timer()
If Shape1.Top > Gravity.Top Then yspeed = yspeed - 1 Else yspeed = yspeed + 1
If Shape1.Left > Gravity.Left Then xspeed = xspeed - 1 Else xspeed = xspeed + 1
Shape1.Top = Shape1.Top + yspeed
Shape1.Left = Shape1.Left + xspeed
Picture1.Line (Shape1.Left, Shape1.Top)-(Shape1.Left, Shape1.Top)
End Sub
Private Function Process(number As Integer) As Integer
Process = number
Select Case number
Case 48 To 57
Case 8
Case Else
Process = 0
End Select
End Function
