VERSION 5.00
Begin VB.Form ProgressBar_1 
   Caption         =   "Progress Bar Demo"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Timer off"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Timer On"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add one"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2010
      TabIndex        =   0
      Top             =   1770
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   495
      Left            =   240
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   390
      Left            =   285
      Top             =   1725
      Width           =   3885
   End
End
Attribute VB_Name = "ProgressBar_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngWidth As Long
Dim strError As String

Dim intValueAdder As Integer

Private Sub BarValue(Value As Integer)
If Value > 100 Then
    strError = "Value over 100"
    Exit Sub
ElseIf Value < 0 Then
    strError = "Value under 0"
    Exit Sub
Else
    Shape2.Width = Value * lngWidth
    Label1.Caption = Value & "%"
End If
End Sub

Private Function LastError(Error As String)
Error = strError
End Function

Private Sub Command1_Click()
intValueAdder = intValueAdder + 1
If intValueAdder > 100 Then intValueAdder = 0
BarValue intValueAdder
End Sub

Private Sub Form_Load()
lngWidth = Shape2.Width / 100
Shape2.Width = 0
Call BarValue(0)
End Sub

Private Sub Option1_Click()
Timer1.Enabled = True
End Sub

Private Sub Option2_Click()
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
intValueAdder = intValueAdder + 1
If intValueAdder > 100 Then intValueAdder = 0
BarValue intValueAdder
End Sub
