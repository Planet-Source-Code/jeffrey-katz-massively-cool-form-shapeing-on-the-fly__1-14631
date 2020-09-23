VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   $"Form1.frx":0000
      Height          =   1095
      Left            =   720
      TabIndex        =   1
      Top             =   900
      Width           =   3075
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Size"
      Height          =   195
      Left            =   4230
      TabIndex        =   0
      Top             =   2880
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
DoEvents 'Show the form
Me.Show 'Show the form
docoolthing 'Open In A Cool Way
End Sub

Sub docoolthing()
DoEvents

a = Me.Width
b = Me.Height
    
Me.Height = 450
Me.Width = 450
    
    For i = 20 To a Step 16
        Me.Width = i
        ReleaseCapture
        ShapeTheForm Me
    Next
    
    For i = 20 To b Step 64
        Me.Height = i
        ReleaseCapture
        ShapeTheForm Me
    Next
    
    'sweet huh?
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF012, 0 'Move the form
End Sub

Private Sub Form_Resize()
Label1.Top = Me.ScaleHeight - Label1.Height
Label1.Left = Me.ScaleWidth - Label1.Width 'move the resize label
ShapeTheForm Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, SIZE_SE, 0 'THIS is importent SIZE_SE is nowhere in winapi
End Sub
